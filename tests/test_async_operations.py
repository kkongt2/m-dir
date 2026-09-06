import os
import subprocess
import sys
import textwrap
import unittest
from pathlib import Path


class AsyncOperationTests(unittest.TestCase):
    def run_gui_check(self, code):
        result = subprocess.run(
            [sys.executable, "-c", textwrap.dedent(code)],
            cwd=Path(__file__).resolve().parents[1],
            env={**os.environ, "QT_QPA_PLATFORM": "offscreen"},
            capture_output=True, text=True, timeout=20,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_file_operations_imports_without_explorer_widgets(self):
        self.run_gui_check('''
            import sys
            import file_operations
            assert 'multipane_explorer' not in sys.modules
            assert 'PyQt5.QtWidgets' not in sys.modules
        ''')

    def test_application_starts_and_closes_with_isolated_settings(self):
        self.run_gui_check('''
            import tempfile, time
            from pathlib import Path
            from PyQt5 import QtCore, QtWidgets
            import multipane_explorer as e
            app = QtWidgets.QApplication([])
            with tempfile.TemporaryDirectory() as root:
                QtCore.QSettings.setDefaultFormat(QtCore.QSettings.IniFormat)
                QtCore.QSettings.setPath(QtCore.QSettings.IniFormat, QtCore.QSettings.UserScope, root)
                QtCore.QSettings.setPath(QtCore.QSettings.IniFormat, QtCore.QSettings.SystemScope, root)
                data = Path(root, 'data')
                data.mkdir()
                (data / 'sample.txt').write_text('sample')
                window = e.MultiExplorer(pane_count=4, start_paths=[str(data)] * 4)
                window.show()
                deadline = time.monotonic() + 8
                while time.monotonic() < deadline:
                    app.processEvents()
                    if all(p._fast_enum_done for p in window.panes):
                        break
                    time.sleep(.005)
                assert len(window.panes) == 4
                assert all(p._fast_enum_done for p in window.panes), 'folder listing did not complete'
                assert window.close(), 'application did not shut down cleanly'
                app.processEvents()
        ''')

    def test_slow_suggestions_keep_ui_responsive_and_discard_old_results(self):
        self.run_gui_check('''
            import threading, time
            from unittest import mock
            from PyQt5 import QtCore, QtWidgets
            import multipane_explorer as e
            import file_operations as operations
            app = QtWidgets.QApplication([])
            e.PathBar._shared_recent_paths = []
            entered, release = threading.Event(), threading.Event()
            ticks = []
            def collect(worker, typed, limit):
                if typed == 'old':
                    entered.set()
                    release.wait(3)
                return [typed + '-result']
            def pump_until(check):
                deadline = time.monotonic() + 4
                while not check() and time.monotonic() < deadline:
                    app.processEvents()
                    time.sleep(.001)
                assert check(), 'timed out'
            with mock.patch.object(e.PathSuggestionWorker, '_collect_filesystem_suggestions', collect):
                bar = e.PathBar()
                try:
                    bar._edit.setText('old')
                    bar._refresh_edit_completer()
                    pump_until(entered.is_set)
                    QtCore.QTimer.singleShot(0, lambda: ticks.append(True))
                    bar._edit.setText('new')
                    bar._refresh_edit_completer()
                    pump_until(lambda: bool(ticks))
                    assert not release.is_set(), 'UI did not run independently'
                    release.set()
                    pump_until(lambda: 'new-result' in bar._edit_model.stringList())
                    assert 'old-result' not in bar._edit_model.stringList()
                finally:
                    release.set()
                    bar.shutdown_suggestions()
                    assert e._cancel_and_wait_child_threads(bar, 4000)
                    app.processEvents()
        ''')

    def test_navigation_and_paste_checks_do_not_block_ui(self):
        self.run_gui_check('''
            import tempfile, threading, time
            from pathlib import Path
            from unittest import mock
            from PyQt5 import QtCore, QtWidgets
            import multipane_explorer as e
            app = QtWidgets.QApplication([])
            def pump(check):
                deadline = time.monotonic() + 4
                while not check() and time.monotonic() < deadline:
                    app.processEvents()
                    time.sleep(.001)
                assert check(), 'timed out'
            with tempfile.TemporaryDirectory() as root:
                # Windows resolves actual path casing (including runner temp paths).
                # Exercise that mismatch instead of relying on local temp casing.
                if e.os.name == 'nt':
                    root = root.swapcase()
                QtCore.QSettings.setDefaultFormat(QtCore.QSettings.IniFormat)
                QtCore.QSettings.setPath(QtCore.QSettings.IniFormat, QtCore.QSettings.UserScope, root)
                old, new = Path(root, 'old'), Path(root, 'new')
                old.mkdir(); new.mkdir()
                old_path, new_path = e.nice_path(str(old)), e.nice_path(str(new))
                window = e.MultiExplorer(pane_count=4, start_paths=[root] * 4)
                pump(lambda: all(p._fast_enum_done for p in window.panes))
                pane = window.panes[0]
                entered, release = threading.Event(), threading.Event()
                original = e.os.path.isdir
                def slow(path):
                    if str(path) == old_path:
                        entered.set(); release.wait(3)
                    return original(path)
                try:
                    with mock.patch.object(e.os.path, 'isdir', side_effect=slow):
                        pane.set_path(str(old))
                        pump(entered.is_set)
                        ticks = []
                        QtCore.QTimer.singleShot(0, lambda: ticks.append(True))
                        pane.set_path(str(new))
                        pump(lambda: ticks and pane.current_path() == new_path)
                        assert not release.is_set()
                        release.set()
                        pump(lambda: all(not w.isRunning() for w in pane.findChildren(e.BackgroundCheck)))
                        app.processEvents()
                        assert pane.current_path() == new_path
                    entered.clear(); release.clear()
                    def prepare(*args):
                        entered.set(); release.wait(3)
                        return ([], [], [], {}, [])
                    with mock.patch.object(e, 'prepare_file_operation', side_effect=prepare), mock.patch.object(pane, '_start_prepared_op') as finish:
                        pane._start_bg_op('copy', [str(old)], str(new))
                        pump(entered.is_set)
                        ticks.clear()
                        QtCore.QTimer.singleShot(0, lambda: ticks.append(True))
                        pump(lambda: bool(ticks))
                        assert not finish.called
                        release.set()
                        pump(lambda: finish.called)
                finally:
                    release.set()
                    assert window.close()
                    app.processEvents()
        ''')

    def test_undo_runs_in_background_and_preserves_cancelled_remainder(self):
        self.run_gui_check('''
            import tempfile, threading, time, types
            from pathlib import Path
            from unittest import mock
            from PyQt5 import QtCore, QtWidgets
            import multipane_explorer as e
            import file_operations as operations
            app = QtWidgets.QApplication([])
            entered, release = threading.Event(), threading.Event()
            shown = []
            class Pane(QtCore.QObject):
                undo_last = e.ExplorerPane.undo_last
                _on_undo_worker_finished = e.ExplorerPane._on_undo_worker_finished
                def window(self): return types.SimpleNamespace(winId=lambda: 0)
                def _show_pane_progress(self, *args, **kwargs): shown.append(kwargs.get('busy'))
                def _set_pane_progress_value(self, *args): pass
                def _set_pane_progress_status(self, *args): pass
                def _hide_pane_progress(self): pass
                def refresh(self): pass
            def recycle(path, hwnd):
                entered.set()
                release.wait(3)
                return True
            def pump_until(check):
                deadline = time.monotonic() + 4
                while not check() and time.monotonic() < deadline:
                    app.processEvents()
                    time.sleep(.001)
                assert check(), 'timed out'
            with tempfile.TemporaryDirectory() as root, mock.patch.object(operations, 'recycle_path_to_trash', recycle):
                paths = [str(Path(root, name)) for name in ('first', 'second')]
                for path in paths: Path(path).write_text('payload')
                pane = Pane()
                pane.host = types.SimpleNamespace(file_ops=e.FileOperationManager(),
                    flash_status=lambda *a: None, show_operation_status=lambda *a: None)
                pane._undo_stack = [{'type': 'remove_created', 'paths': list(paths)}]
                pane._undo_worker = pane._file_worker = None
                pane.undo_last()
                worker = pane._undo_worker
                try:
                    pump_until(entered.is_set)
                    pump_until(lambda: False in shown)
                    QtCore.QTimer.singleShot(0, worker.cancel)
                    pump_until(lambda: worker._cancel)
                    release.set()
                    pump_until(lambda: pane._undo_worker is None)
                    assert pane._undo_stack[0]['paths'] == [paths[0]]
                finally:
                    release.set()
                    assert pane.host.file_ops.cancel_all(4000)
                    app.processEvents()
        ''')


if __name__ == "__main__":
    unittest.main()
