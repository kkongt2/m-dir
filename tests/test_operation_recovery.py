import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import multipane_explorer as explorer


class MoveSourceProtectionTests(unittest.TestCase):
    def test_changes_during_copy_preserve_source(self):
        for change in ("add", "modify", "replace"):
            with self.subTest(change=change), tempfile.TemporaryDirectory() as root:
                source = Path(root, "source")
                source.mkdir()
                original = source / "old.txt"
                original.write_text("original")
                destination = Path(root, "destination")
                destination.mkdir()
                worker = explorer.FileOpWorker("move", [str(source)], str(destination))
                copy = worker._copy_to_new_path

                def copy_then_change(src, dst):
                    ok = copy(src, dst)
                    if change == "add":
                        (source / "new.txt").write_text("new data")
                    elif change == "modify":
                        original.write_text("changed data")
                    else:
                        original.rename(source / "saved.txt")
                        original.write_text("replacement data")
                    return ok

                with mock.patch.object(explorer, "_same_filesystem", return_value=False), mock.patch.object(
                    worker, "_copy_to_new_path", side_effect=copy_then_change
                ):
                    self.assertFalse(worker._move_source_transactional(
                        str(source), str(destination / "source"), None, False))
                self.assertTrue(original.exists())
                self.assertEqual((destination / "source" / "old.txt").read_text(), "original")
                if change == "add":
                    self.assertEqual((source / "new.txt").read_text(), "new data")
                self.assertGreater(worker.error_count, 0)
                self.assertFalse(worker.undo_move_pairs)

    def test_new_file_during_cleanup_is_not_deleted(self):
        with tempfile.TemporaryDirectory() as root:
            source = Path(root, "source")
            source.mkdir()
            original = source / "old.txt"
            original.write_text("copied")
            snapshot = explorer._snapshot_move_source(str(source))
            remove = explorer.remove_any

            def remove_then_add(path):
                remove(path)
                (source / "new.txt").write_text("keep")

            with mock.patch.object(explorer, "remove_any", side_effect=remove_then_add):
                errors = explorer._cleanup_copied_source(str(source), snapshot)
            self.assertTrue(errors)
            self.assertEqual((source / "new.txt").read_text(), "keep")


class DestinationProtectionTests(unittest.TestCase):
    def test_failed_copy_does_not_remove_unowned_destination(self):
        with tempfile.TemporaryDirectory() as root:
            source = Path(root, "source.txt")
            source.write_text("source")
            destination = Path(root, "destination.txt")
            destination.write_text("other process")
            worker = explorer.FileOpWorker("copy", [str(source)], root)
            with mock.patch.object(worker, "_copy_to_new_path", side_effect=OSError("copy failed")):
                self.assertFalse(worker._copy_source_transactional(str(source), str(destination), None, False))
            self.assertEqual(destination.read_text(), "other process")

    def test_concurrent_destination_is_not_overwritten(self):
        for operation in ("copy", "move"):
            with self.subTest(operation=operation), tempfile.TemporaryDirectory() as root:
                source = Path(root, "source.txt")
                source.write_text("source")
                destination = Path(root, "destination.txt")
                worker = explorer.FileOpWorker(operation, [str(source)], root)
                copy = worker._copy_to_new_path

                def copy_then_create(src, dst):
                    result = copy(src, dst)
                    destination.write_text("other process")
                    return result

                method = getattr(worker, f"_{operation}_source_transactional")
                with mock.patch.object(explorer, "_same_filesystem", return_value=False), mock.patch.object(
                    worker, "_copy_to_new_path", side_effect=copy_then_create
                ):
                    self.assertFalse(method(str(source), str(destination), None, False))
                self.assertEqual(destination.read_text(), "other process")
                self.assertEqual(source.read_text(), "source")
                self.assertEqual(sorted(p.name for p in Path(root).iterdir()), ["destination.txt", "source.txt"])

    def test_overwrite_failure_preserves_original_backup_and_concurrent_file(self):
        with tempfile.TemporaryDirectory() as root:
            source = Path(root, "source.txt")
            source.write_text("source")
            destination = Path(root, "destination.txt")
            destination.write_text("original")
            worker = explorer.FileOpWorker("copy", [str(source)], root)
            backup = worker._backup_destination

            def backup_then_create(*args):
                saved = backup(*args)
                destination.write_text("other process")
                return saved

            with mock.patch.object(worker, "_backup_destination", side_effect=backup_then_create):
                self.assertFalse(worker._copy_source_transactional(str(source), str(destination), "overwrite", True))
            self.assertEqual(destination.read_text(), "other process")
            backups = list(Path(root).glob(".__mprn_backup_*"))
            self.assertEqual(len(backups), 1)
            self.assertEqual(backups[0].read_text(), "original")
            self.assertTrue(any(str(backups[0]) in error for error in worker.errors))


if __name__ == "__main__":
    unittest.main()
