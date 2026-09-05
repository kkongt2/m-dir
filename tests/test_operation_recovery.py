import os
import tempfile
import types
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


class UndoRecoveryTests(unittest.TestCase):
    def test_queued_undo_cancellation_does_not_touch_files(self):
        action = {"type": "remove_created", "paths": ["untouched"]}
        worker = explorer.UndoWorker(action)
        pane = types.SimpleNamespace(_file_worker=worker, btn_op_cancel=mock.Mock(),
                                     _set_pane_progress_status=mock.Mock())
        explorer.ExplorerPane._request_file_op_cancel(pane)
        with mock.patch.object(explorer, "recycle_path_to_trash") as recycle:
            worker.run()
        recycle.assert_not_called()
        self.assertEqual(worker.remaining_action, action)
        self.assertEqual(worker.failure_message, "Operation cancelled.")

    def test_rename_group_undo_restores_overlapping_names(self):
        with tempfile.TemporaryDirectory() as root:
            a, b, c = [Path(root, n) for n in ("a", "b", "c")]
            a.write_text("a")
            b.write_text("b")
            pairs = explorer.execute_bulk_rename_transaction([(str(a), str(b)), (str(b), str(c))])
            worker = explorer.UndoWorker({"type": "move_back", "pairs": pairs, "rename_group": True})
            worker.run()
            self.assertTrue(worker.completed, worker.failure_message)
            self.assertEqual((a.read_text(), b.read_text()), ("a", "b"))
            self.assertFalse(c.exists())

    def test_cancelled_copy_registers_completed_items_once(self):
        with tempfile.TemporaryDirectory() as root:
            source = Path(root, "source.txt")
            source.write_text("source")
            target = Path(root, "target")
            target.mkdir()
            worker = explorer.FileOpWorker("copy", [str(source)], str(target))
            copy = worker._copy_source_transactional
            errors = []
            worker.error.connect(errors.append)

            def copy_then_cancel(*args):
                ok = copy(*args)
                worker.cancel()
                return ok

            with mock.patch.object(worker, "_copy_source_transactional", side_effect=copy_then_cancel):
                worker.run()
            pane = types.SimpleNamespace(_undo_stack=[])
            explorer.ExplorerPane._push_file_op_undo(pane, worker, "copy")
            explorer.ExplorerPane._push_file_op_undo(pane, worker, "copy")
            self.assertEqual(errors, ["Operation cancelled."])
            self.assertEqual(len(pane._undo_stack), 1)
            self.assertEqual(pane._undo_stack[0]["paths"], [str(target / source.name)])

    def test_failed_undo_retains_only_unfinished_paths(self):
        with tempfile.TemporaryDirectory() as root:
            paths = [str(Path(root, name)) for name in ("fail.txt", "ok.txt")]
            for path in paths:
                Path(path).write_text("payload")
            action = {"type": "remove_created", "paths": list(paths)}
            worker = explorer.UndoWorker(action)
            with mock.patch.object(explorer, "recycle_path_to_trash", side_effect=lambda p, _h: p == paths[1]):
                worker.run()
            self.assertFalse(worker.completed)
            self.assertEqual(worker.remaining_action, {"type": "remove_created", "paths": [paths[0]]})
            self.assertEqual(action["paths"], paths, "worker must not mutate the UI-owned record")



class BulkRenameRecoveryTests(unittest.TestCase):
    def test_failure_restores_chains_and_cycles(self):
        for cycle in (False, True):
            with self.subTest(cycle=cycle), tempfile.TemporaryDirectory() as root:
                paths = [Path(root, f"{i}.txt") for i in range(1, 5)]
                for p in paths[:3]:
                    p.write_text(p.stem)
                targets = [paths[1], paths[0] if cycle else paths[2], paths[3]]
                operations = [(str(src), str(dst)) for src, dst in zip(paths, targets)]
                rename = explorer._rename_no_replace
                calls = 0

                def fail_last_commit(src, dst):
                    nonlocal calls
                    calls += 1
                    if calls == 6:
                        raise OSError("commit failed")
                    rename(src, dst)

                with mock.patch.object(explorer, "_rename_no_replace", side_effect=fail_last_commit):
                    with self.assertRaisesRegex(OSError, "commit failed"):
                        explorer.execute_bulk_rename_transaction(operations)
                self.assertEqual({p.name: p.read_text() for p in Path(root).iterdir()},
                                 {f"{i}.txt": str(i) for i in range(1, 4)})

    def test_successful_cycle_swaps_contents_without_temporary_files(self):
        with tempfile.TemporaryDirectory() as root:
            a, b = Path(root, "a.txt"), Path(root, "b.txt")
            a.write_text("a")
            b.write_text("b")
            explorer.execute_bulk_rename_transaction([(str(a), str(b)), (str(b), str(a))])
            self.assertEqual((a.read_text(), b.read_text()), ("b", "a"))
            self.assertEqual(len(list(Path(root).iterdir())), 2)


class DestinationProtectionTests(unittest.TestCase):
    def test_successful_overwrite_removes_only_the_previous_destination(self):
        for operation in ("copy", "move"):
            for directory in (False, True):
                with self.subTest(operation=operation, directory=directory), tempfile.TemporaryDirectory() as root:
                    source, destination = Path(root, "source"), Path(root, "destination")
                    if directory:
                        source.mkdir()
                        destination.mkdir()
                        (source / "new.txt").write_text("new")
                        (destination / "old.txt").write_text("old")
                    else:
                        source.write_text("new")
                        destination.write_text("old")
                    worker = explorer.FileOpWorker(operation, [str(source)], root)
                    method = getattr(worker, f"_{operation}_source_transactional")
                    with mock.patch.object(explorer, "_same_filesystem", return_value=False):
                        self.assertTrue(method(str(source), str(destination), "overwrite", True), worker.errors)
                    self.assertEqual((destination / "new.txt" if directory else destination).read_text(), "new")
                    self.assertEqual(source.exists(), operation == "copy")
                    self.assertFalse(list(Path(root).glob(".__mprn_*")))

    def test_overwrite_cancellation_preserves_both_originals(self):
        for operation in ("copy", "move"):
            with self.subTest(operation=operation), tempfile.TemporaryDirectory() as root:
                source = Path(root, "source.txt")
                source.write_text("source")
                destination = Path(root, "destination.txt")
                destination.write_text("destination")
                worker = explorer.FileOpWorker(operation, [str(source)], root)
                copy = worker._copy_to_new_path

                def copy_then_cancel(*args):
                    ok = copy(*args)
                    worker.cancel()
                    return ok

                method = getattr(worker, f"_{operation}_source_transactional")
                with mock.patch.object(explorer, "_same_filesystem", return_value=False), mock.patch.object(
                    worker, "_copy_to_new_path", side_effect=copy_then_cancel
                ):
                    self.assertFalse(method(str(source), str(destination), "overwrite", True))
                self.assertEqual(source.read_text(), "source")
                self.assertEqual(destination.read_text(), "destination")
                self.assertFalse(list(Path(root).glob(".__mprn_*")))

    def test_destination_change_during_overwrite_is_preserved(self):
        with tempfile.TemporaryDirectory() as root:
            source, destination = Path(root, "source"), Path(root, "destination")
            source.write_text("source")
            destination.write_text("original")
            worker = explorer.FileOpWorker("copy", [str(source)], root)
            copy = worker._copy_to_new_path

            def copy_then_modify(*args):
                result = copy(*args)
                destination.write_text("updated by another program")
                return result

            with mock.patch.object(worker, "_copy_to_new_path", side_effect=copy_then_modify):
                self.assertFalse(worker._copy_source_transactional(str(source), str(destination), "overwrite", True))
            self.assertEqual(destination.read_text(), "updated by another program")
            self.assertFalse(list(Path(root).glob(".__mprn_*")))

    def test_failed_atomic_move_never_removes_existing_destination(self):
        with tempfile.TemporaryDirectory() as root:
            source, destination = Path(root, "source"), Path(root, "destination")
            source.write_text("source")
            destination.write_text("concurrent file")
            worker = explorer.FileOpWorker("move", [str(source)], root)
            self.assertFalse(worker._move_source_transactional(str(source), str(destination), None, False))
            self.assertEqual(source.read_text(), "source")
            self.assertEqual(destination.read_text(), "concurrent file")

    def test_overwrite_promotion_failure_restores_original(self):
        for operation in ("copy", "move"):
            with self.subTest(operation=operation), tempfile.TemporaryDirectory() as root:
                source, destination = Path(root, "source"), Path(root, "destination")
                source.write_text("source")
                destination.write_text("original")
                worker = explorer.FileOpWorker(operation, [str(source)], root)
                rename = explorer._rename_no_replace

                def fail_promotion(src, dst):
                    if Path(src).name == "payload" and dst == str(destination):
                        raise PermissionError("destination locked")
                    return rename(src, dst)

                method = getattr(worker, f"_{operation}_source_transactional")
                with mock.patch.object(explorer, "_same_filesystem", return_value=False), mock.patch.object(
                    explorer, "_rename_no_replace", side_effect=fail_promotion
                ):
                    self.assertFalse(method(str(source), str(destination), "overwrite", True))
                self.assertEqual(source.read_text(), "source")
                self.assertEqual(destination.read_text(), "original")
                self.assertFalse(list(Path(root).glob(".__mprn_*")))

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
