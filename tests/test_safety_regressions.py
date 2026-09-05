import os
import errno
import subprocess
import tempfile
import time
import types
import unittest
from pathlib import Path
from unittest import mock

from PyQt5 import QtCore

import multipane_explorer as explorer


class _FakeSettings:
    def __init__(self, value):
        self._value = value

    def value(self, _key, default=None):
        return self._value if self._value is not None else default


class FileOperationSafetyTests(unittest.TestCase):
    def _make_source_tree(self, root: str) -> str:
        src = os.path.join(root, "source")
        os.makedirs(src)
        Path(src, "first.txt").write_text("first", encoding="utf-8")
        Path(src, "second.txt").write_text("second", encoding="utf-8")
        return src

    def test_normal_copy_does_not_force_per_file_fsync(self):
        with tempfile.TemporaryDirectory() as root:
            src = os.path.join(root, "source.txt")
            dst = os.path.join(root, "destination.txt")
            Path(src).write_bytes(b"payload")
            worker = explorer.FileOpWorker("copy", [src], root, durable_copies=False)

            with mock.patch.object(explorer.os, "fsync") as fsync:
                self.assertTrue(worker._copy_file(src, dst))

            fsync.assert_not_called()
            self.assertEqual(Path(dst).read_bytes(), b"payload")

    def test_durable_copy_keeps_explicit_fsync_option(self):
        with tempfile.TemporaryDirectory() as root:
            src = os.path.join(root, "source.txt")
            dst = os.path.join(root, "destination.txt")
            Path(src).write_bytes(b"payload")
            worker = explorer.FileOpWorker("copy", [src], root, durable_copies=True)

            with mock.patch.object(explorer.os, "fsync") as fsync:
                self.assertTrue(worker._copy_file(src, dst))

            fsync.assert_called_once()

    def test_cross_filesystem_cleanup_failure_keeps_complete_destination(self):
        with tempfile.TemporaryDirectory() as root:
            src = self._make_source_tree(root)
            dst_dir = os.path.join(root, "destination")
            os.makedirs(dst_dir)
            dst = os.path.join(dst_dir, "source")
            worker = explorer.FileOpWorker("move", [src], dst_dir)

            def partial_source_cleanup(path, **_kwargs):
                os.remove(os.path.join(path, "second.txt"))
                return 1, ["simulated source cleanup failure"]

            with mock.patch.object(explorer, "_same_filesystem", return_value=False), mock.patch.object(
                explorer, "delete_any_permanent_best_effort", side_effect=partial_source_cleanup
            ):
                result = worker._move_source_transactional(src, dst, None, False)

            self.assertFalse(result)
            self.assertTrue(os.path.exists(os.path.join(src, "first.txt")))
            self.assertEqual(
                sorted(os.listdir(dst)),
                ["first.txt", "second.txt"],
                "the verified destination must remain complete when source cleanup is partial",
            )
            self.assertGreater(worker.error_count, 0)
            self.assertEqual(worker.undo_move_pairs, [])

    def test_cross_filesystem_move_success_removes_source_after_promotion(self):
        with tempfile.TemporaryDirectory() as root:
            src = self._make_source_tree(root)
            dst_dir = os.path.join(root, "destination")
            os.makedirs(dst_dir)
            dst = os.path.join(dst_dir, "source")
            worker = explorer.FileOpWorker("move", [src], dst_dir)

            with mock.patch.object(explorer, "_same_filesystem", return_value=False):
                result = worker._move_source_transactional(src, dst, None, False)

            self.assertTrue(result)
            self.assertFalse(os.path.lexists(src))
            self.assertEqual(sorted(os.listdir(dst)), ["first.txt", "second.txt"])
            self.assertEqual(worker.error_count, 0)

    def test_same_filesystem_move_uses_atomic_replace(self):
        with tempfile.TemporaryDirectory() as root:
            src = os.path.join(root, "source.txt")
            Path(src).write_text("payload", encoding="utf-8")
            dst_dir = os.path.join(root, "destination")
            os.makedirs(dst_dir)
            dst = os.path.join(dst_dir, "source.txt")
            worker = explorer.FileOpWorker("move", [src], dst_dir)

            with mock.patch.object(explorer, "_same_filesystem", return_value=True), mock.patch.object(
                explorer.shutil, "move", side_effect=AssertionError("shutil.move must not be used")
            ):
                result = worker._move_source_transactional(src, dst, None, False)

            self.assertTrue(result)
            self.assertFalse(os.path.lexists(src))
            self.assertEqual(Path(dst).read_text(encoding="utf-8"), "payload")

    def test_cross_device_rename_error_switches_to_safe_move(self):
        with tempfile.TemporaryDirectory() as root:
            src = os.path.join(root, "source.txt")
            Path(src).write_text("payload", encoding="utf-8")
            dst_dir = os.path.join(root, "destination")
            os.makedirs(dst_dir)
            dst = os.path.join(dst_dir, "source.txt")
            worker = explorer.FileOpWorker("move", [src], dst_dir)
            real_replace = os.replace

            def replace_with_cross_device_detection(source, destination):
                if source == src and destination == dst:
                    raise OSError(errno.EXDEV, "simulated filesystem boundary")
                return real_replace(source, destination)

            with mock.patch.object(explorer, "_same_filesystem", return_value=True), mock.patch.object(
                explorer.os, "replace", side_effect=replace_with_cross_device_detection
            ):
                result = worker._move_source_transactional(src, dst, None, False)

            self.assertTrue(result)
            self.assertFalse(os.path.lexists(src))
            self.assertEqual(Path(dst).read_text(encoding="utf-8"), "payload")

    def test_cross_filesystem_overwrite_cleanup_failure_keeps_new_destination(self):
        with tempfile.TemporaryDirectory() as root:
            src = self._make_source_tree(root)
            dst_dir = os.path.join(root, "destination")
            dst = os.path.join(dst_dir, "source")
            os.makedirs(dst)
            Path(dst, "old.txt").write_text("old", encoding="utf-8")
            worker = explorer.FileOpWorker("move", [src], dst_dir, conflict_map={src: "overwrite"})

            def partial_source_cleanup(path, **_kwargs):
                os.remove(os.path.join(path, "second.txt"))
                return 1, ["simulated source cleanup failure"]

            with mock.patch.object(explorer, "_same_filesystem", return_value=False), mock.patch.object(
                explorer, "delete_any_permanent_best_effort", side_effect=partial_source_cleanup
            ):
                result = worker._move_source_transactional(src, dst, "overwrite", True)

            self.assertFalse(result)
            self.assertEqual(sorted(os.listdir(dst)), ["first.txt", "second.txt"])
            self.assertFalse(any(name.startswith(".__mprn_backup_") for name in os.listdir(dst_dir)))

    def test_copy_rejects_junction_instead_of_traversing_it(self):
        with tempfile.TemporaryDirectory() as root:
            src = self._make_source_tree(root)
            dst_dir = os.path.join(root, "destination")
            os.makedirs(dst_dir)
            dst = os.path.join(dst_dir, "source")
            worker = explorer.FileOpWorker("copy", [src], dst_dir)

            with mock.patch.object(explorer, "_is_junction", side_effect=lambda path: path == src):
                result = worker._copy_source_transactional(src, dst, None, False)

            self.assertFalse(result)
            self.assertTrue(os.path.isdir(src))
            self.assertFalse(os.path.lexists(dst))
            self.assertGreater(worker.error_count, 0)

    @unittest.skipUnless(os.name == "nt", "Windows junction fallback")
    def test_junction_detection_falls_back_to_reparse_tag(self):
        fake_stat = types.SimpleNamespace(st_reparse_tag=0xA0000003)
        with mock.patch.object(explorer.os.path, "isjunction", return_value=False), mock.patch.object(
            explorer.os, "lstat", return_value=fake_stat
        ):
            self.assertTrue(explorer._is_junction(r"C:\\junction"))

    @unittest.skipUnless(os.name == "nt", "Windows junction behavior")
    def test_remove_junction_does_not_delete_its_target(self):
        with tempfile.TemporaryDirectory() as root:
            target = os.path.join(root, "target")
            junction = os.path.join(root, "junction")
            os.makedirs(target)
            marker = Path(target, "marker.txt")
            marker.write_text("keep", encoding="utf-8")
            created = subprocess.run(
                ["cmd.exe", "/c", "mklink", "/J", junction, target],
                capture_output=True,
                text=True,
                check=False,
            )
            if created.returncode != 0:
                self.skipTest(f"could not create junction: {created.stderr or created.stdout}")

            self.assertTrue(explorer._is_junction(junction))
            explorer.remove_any(junction)
            self.assertFalse(os.path.lexists(junction))
            self.assertEqual(marker.read_text(encoding="utf-8"), "keep")

    def test_symbolic_link_copy_preserves_the_link_when_supported(self):
        with tempfile.TemporaryDirectory() as root:
            target = os.path.join(root, "target.txt")
            link = os.path.join(root, "source-link")
            Path(target).write_text("target", encoding="utf-8")
            try:
                os.symlink(target, link)
            except OSError as exc:
                self.skipTest(f"symbolic links are unavailable: {exc}")

            dst_dir = os.path.join(root, "destination")
            os.makedirs(dst_dir)
            dst = os.path.join(dst_dir, "source-link")
            worker = explorer.FileOpWorker("copy", [link], dst_dir)
            result = worker._copy_source_transactional(link, dst, None, False)

            self.assertTrue(result)
            self.assertTrue(os.path.islink(dst))
            self.assertEqual(os.readlink(dst), target)


class PaneRestoreTests(unittest.TestCase):
    def test_cli_pane_count_overrides_saved_layout(self):
        self.assertEqual(explorer._resolve_pane_count(8, _FakeSettings(4)), 8)

    def test_saved_pane_count_is_restored_without_cli_override(self):
        self.assertEqual(explorer._resolve_pane_count(None, _FakeSettings(4)), 4)

    def test_invalid_saved_pane_count_falls_back_to_six(self):
        self.assertEqual(explorer._resolve_pane_count(None, _FakeSettings("invalid")), 6)

    def test_panes_argument_defaults_to_restore_mode(self):
        self.assertIsNone(explorer.parse_args([]).panes)


class ThreadLifecycleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtCore.QCoreApplication.instance() or QtCore.QCoreApplication([])

    def test_child_thread_timeout_prevents_owner_destruction(self):
        class StubbornThread(QtCore.QThread):
            def cancel(self):
                pass

            def run(self):
                time.sleep(0.15)

        owner = QtCore.QObject()
        worker = StubbornThread(owner)
        worker.start()
        self.assertFalse(explorer._cancel_and_wait_child_threads(owner, 10))
        self.assertTrue(worker.wait(1000))

    def test_cooperative_child_thread_stops_within_deadline(self):
        class CooperativeThread(QtCore.QThread):
            def __init__(self, parent=None):
                super().__init__(parent)
                self.cancelled = False

            def cancel(self):
                self.cancelled = True

            def run(self):
                while not self.cancelled:
                    time.sleep(0.005)

        owner = QtCore.QObject()
        worker = CooperativeThread(owner)
        worker.start()
        self.assertTrue(explorer._cancel_and_wait_child_threads(owner, 1000))


if __name__ == "__main__":
    unittest.main()
