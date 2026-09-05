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


if __name__ == "__main__":
    unittest.main()
