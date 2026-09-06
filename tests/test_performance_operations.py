import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import file_operations as operations


class OperationPerformanceTests(unittest.TestCase):
    def test_same_filesystem_move_never_scans_descendants_for_progress(self):
        with tempfile.TemporaryDirectory() as root:
            source = Path(root, "source")
            destination = Path(root, "destination")
            source.mkdir()
            destination.mkdir()
            Path(source, "payload.txt").write_text("preserved", encoding="utf-8")
            worker = operations.FileOpWorker("move", [str(source)], str(destination))
            errors = []
            worker.error.connect(errors.append)
            with mock.patch.object(worker, "_scan_source_progress", side_effect=AssertionError("unnecessary scan")):
                worker.run()
            self.assertEqual(errors, [])
            self.assertFalse(source.exists())
            self.assertEqual(Path(destination, "source", "payload.txt").read_text(), "preserved")
            self.assertEqual(worker._total_items, 1)

    def test_cross_filesystem_progress_still_uses_copy_estimate(self):
        worker = operations.FileOpWorker("move", ["source"], "destination")
        with mock.patch.object(operations, "_same_filesystem", return_value=False), mock.patch.object(
            worker, "_scan_source_progress", return_value=(123, 4)
        ) as scan:
            worker._calc_total()
        scan.assert_called_once_with("source")
        self.assertEqual((worker._total_bytes, worker._total_items), (123, 4))


if __name__ == "__main__":
    unittest.main()
