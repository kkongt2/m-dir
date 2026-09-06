import os
import tempfile
import unittest
from contextlib import contextmanager
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

    def test_directory_copy_starts_before_enumeration_finishes(self):
        with tempfile.TemporaryDirectory() as root:
            source, destination = Path(root, "source"), Path(root, "destination")
            source.mkdir(); destination.mkdir()
            for i in range(3):
                Path(source, f"{i}.txt").write_text(f"payload {i}")
            worker = operations.FileOpWorker("copy", [str(source)], str(destination))
            worker._calc_total()
            real_scan = os.scandir
            copied, closed = [], []
            real_copy = worker._copy_file
            def copy(src, dst):
                result = real_copy(src, dst)
                copied.append(src)
                return result
            @contextmanager
            def incremental(path):
                with real_scan(path) as entries:
                    def stream():
                        for index, entry in enumerate(entries):
                            if index:
                                self.assertTrue(copied, "enumerated entire directory before copying")
                            yield entry
                    try:
                        yield stream()
                    finally:
                        closed.append(True)
            with mock.patch.object(operations.os, "scandir", incremental), mock.patch.object(worker, "_copy_file", side_effect=copy):
                self.assertTrue(worker._copy_dir_recursive(str(source), str(destination)))
            self.assertEqual(len(copied), 3)
            self.assertEqual(closed, [True])
            for i in range(3):
                self.assertEqual(Path(destination, f"{i}.txt").read_text(), f"payload {i}")

    def test_directory_copy_cancellation_closes_iterator_without_draining_it(self):
        with tempfile.TemporaryDirectory() as root:
            source, destination = Path(root, "source"), Path(root, "destination")
            source.mkdir(); destination.mkdir()
            for i in range(10):
                Path(source, f"{i}.txt").touch()
            worker = operations.FileOpWorker("copy", [str(source)], str(destination))
            worker._calc_total()
            real_scan = os.scandir
            enumerated, closed = [], []
            @contextmanager
            def incremental(path):
                with real_scan(path) as entries:
                    def stream():
                        for entry in entries:
                            enumerated.append(entry.name)
                            yield entry
                    try:
                        yield stream()
                    finally:
                        closed.append(True)
            def cancel(*args):
                worker.cancel()
                return False
            with mock.patch.object(operations.os, "scandir", incremental), mock.patch.object(worker, "_copy_file", side_effect=cancel):
                self.assertFalse(worker._copy_dir_recursive(str(source), str(destination)))
            self.assertEqual(len(enumerated), 1)
            self.assertEqual(closed, [True])


if __name__ == "__main__":
    unittest.main()
