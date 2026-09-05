import os
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest import mock

from PyQt5 import QtCore

import multipane_explorer as explorer


class SharedWorkTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtCore.QCoreApplication.instance() or QtCore.QCoreApplication([])

    def test_same_directory_scan_is_single_flight(self):
        cache = explorer.DirectorySnapshotCache(ttl_s=10, max_entries=4)
        real_scandir = os.scandir
        calls = []
        calls_lock = threading.Lock()

        def slow_scandir(path):
            with calls_lock:
                calls.append(path)
            time.sleep(0.05)
            return real_scandir(path)

        with tempfile.TemporaryDirectory() as root:
            for i in range(10):
                Path(root, f"{i}.txt").write_text("x", encoding="utf-8")
            first = explorer.DirEnumWorker(root)
            second = explorer.DirEnumWorker(root)
            with mock.patch.object(explorer, "GLOBAL_DIR_SNAPSHOTS", cache), mock.patch.object(
                explorer.os, "scandir", side_effect=slow_scandir
            ):
                first.start()
                second.start()
                self.assertTrue(first.wait(3000))
                self.assertTrue(second.wait(3000))

        self.assertEqual(len(calls), 1)

    def test_invalidation_rejects_an_old_inflight_snapshot(self):
        cache = explorer.DirectorySnapshotCache(ttl_s=10, max_entries=4)
        state, key, generation, event, _rows, _error = cache.acquire("C:/same", False, False)
        self.assertEqual(state, "leader")
        cache.invalidate("C:/same")
        cache.publish(key, generation, event, [{"name": "stale"}])

        next_state, _key, _generation, _event, _rows, _error = cache.acquire("C:/same", False, False)
        self.assertEqual(next_state, "leader")

    def test_icon_broker_deduplicates_pending_keys(self):
        broker = explorer.ShellIconBroker()
        job = ("ext:.shared-test", "C:/one.shared-test", False)
        explorer.GLOBAL_SHELL_ICON_CACHE.pop(job[0], None)
        explorer.GLOBAL_SHELL_ICON_FAILURES.pop(job[0], None)

        with mock.patch.object(broker, "_start_next") as start_next:
            broker.request([job])
            broker.request([job])

        self.assertEqual(broker._queue, [job])
        self.assertEqual(broker._pending, {job[0]})
        start_next.assert_called_once()


if __name__ == "__main__":
    unittest.main()
