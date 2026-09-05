import os
import tempfile
import threading
import time
import types
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

    def test_duplicate_watcher_invalidations_are_coalesced(self):
        cache = explorer.DirectorySnapshotCache(ttl_s=10, max_entries=4)
        self.assertTrue(cache.invalidate("C:/same", coalesce_s=1.0))
        self.assertFalse(cache.invalidate("C:/same", coalesce_s=1.0))

    def test_file_changes_wait_for_managed_operations(self):
        state = explorer.FileChangeRefreshState()
        self.assertFalse(state.note_change(operation_pending=True))
        self.assertTrue(state.pending)
        self.assertFalse(state.operation_state_changed(busy=True))
        self.assertTrue(state.operation_state_changed(busy=False))
        self.assertFalse(state.pending)

    def test_header_width_settings_are_deferred(self):
        pane = types.SimpleNamespace(
            _pending_search_widths={},
            _schedule_ui_settings_sync=mock.Mock(),
        )

        explorer.ExplorerPane._save_search_header_width(pane, 3, 144)

        self.assertEqual(pane._pending_search_widths, {3: 144})
        pane._schedule_ui_settings_sync.assert_called_once()

    def test_deferred_settings_are_flushed_together(self):
        settings = mock.Mock()
        pane = types.SimpleNamespace(
            pane_id=2,
            _sort_column=3,
            _sort_order=QtCore.Qt.DescendingOrder,
            _pending_sort_settings=True,
            _pending_search_widths={1: 80, 3: 140},
        )

        with mock.patch.object(explorer, "QSettings", return_value=settings):
            explorer.ExplorerPane._flush_pending_ui_settings(pane)

        self.assertEqual(settings.setValue.call_count, 4)
        settings.sync.assert_called_once()
        self.assertFalse(pane._pending_sort_settings)
        self.assertEqual(pane._pending_search_widths, {})

    def test_shell_icon_api_returns_the_cached_binding(self):
        sentinel = object()
        with mock.patch.object(explorer.sys, "platform", "win32"), mock.patch.object(
            explorer, "_SHELL_ICON_API_INITIALIZED", True
        ), mock.patch.object(explorer, "_SHELL_ICON_API", sentinel):
            self.assertIs(explorer._get_windows_shell_icon_api(), sentinel)
            self.assertIs(explorer._get_windows_shell_icon_api(), sentinel)


if __name__ == "__main__":
    unittest.main()
