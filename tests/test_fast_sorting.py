import unittest
import os
import tempfile
from pathlib import Path

from PyQt5 import QtCore

import multipane_explorer as explorer


class FastRecordSortingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtCore.QCoreApplication.instance() or QtCore.QCoreApplication([])

    def _model(self):
        model = explorer.FastDirModel()
        model.append_rows([
            {"name": "z-dir", "name_l": "z-dir", "path": "C:/z-dir", "is_dir": True, "ext": "", "size": 0, "mtime": 4, "icon_key": "folder"},
            {"name": "z.txt", "name_l": "z.txt", "path": "C:/z.txt", "is_dir": False, "ext": "txt", "size": 30, "mtime": 3, "icon_key": "ext:.txt"},
            {"name": "a-dir", "name_l": "a-dir", "path": "C:/a-dir", "is_dir": True, "ext": "", "size": 0, "mtime": 2, "icon_key": "folder"},
            {"name": "a.txt", "name_l": "a.txt", "path": "C:/a.txt", "is_dir": False, "ext": "txt", "size": 10, "mtime": 1, "icon_key": "ext:.txt"},
        ])
        proxy = explorer.RecordSortProxy()
        proxy.setSourceModel(model)
        return model, proxy

    @staticmethod
    def _names(proxy):
        return [proxy.index(row, 0).data() for row in range(proxy.rowCount())]

    def test_name_sort_keeps_directories_first_in_both_directions(self):
        _model, proxy = self._model()
        proxy.sort(0, QtCore.Qt.AscendingOrder)
        self.assertEqual(self._names(proxy), ["a-dir", "z-dir", "a.txt", "z.txt"])
        proxy.sort(0, QtCore.Qt.DescendingOrder)
        self.assertEqual(self._names(proxy), ["z-dir", "a-dir", "z.txt", "a.txt"])

    def test_numeric_sort_uses_record_keys(self):
        _model, proxy = self._model()
        proxy.sort(1, QtCore.Qt.AscendingOrder)
        self.assertEqual(self._names(proxy), ["a-dir", "z-dir", "a.txt", "z.txt"])

    def test_persistent_index_tracks_the_same_path_after_sort(self):
        model, proxy = self._model()
        persistent = QtCore.QPersistentModelIndex(model.index(0, 0))
        proxy.sort(0, QtCore.Qt.AscendingOrder)
        self.assertTrue(persistent.isValid())
        self.assertEqual(persistent.data(QtCore.Qt.UserRole), "C:/z-dir")

    def test_stat_updates_are_applied_in_consolidated_ranges(self):
        model = explorer.FastDirModel()
        rows = [
            {
                "name": f"{i}.txt",
                "name_l": f"{i}.txt",
                "path": f"C:/{i}.txt",
                "is_dir": False,
                "ext": "txt",
                "size": None,
                "mtime": None,
                "icon_key": "ext:.txt",
            }
            for i in range(100)
        ]
        model.append_rows(rows)
        emissions = []
        model.dataChanged.connect(lambda top, bottom, _roles: emissions.append((top.column(), top.row(), bottom.row())))

        model.apply_stat_batch([(f"C:/{i}.txt", i, 1000 + i) for i in range(100)])

        self.assertEqual(emissions, [(1, 0, 99), (3, 0, 99)])
        self.assertEqual(model.index(99, 1).data(QtCore.Qt.EditRole), 99)

    def test_fast_stat_worker_emits_bounded_batches(self):
        with tempfile.TemporaryDirectory() as root:
            paths = []
            for i in range(130):
                path = os.path.join(root, f"{i}.txt")
                Path(path).write_text("x", encoding="utf-8")
                paths.append(path)
            worker = explorer.FastStatWorker(root, paths)
            batches = []
            worker.statBatchReady.connect(lambda batch: batches.append(batch))

            worker.run()

        self.assertEqual([len(batch) for batch in batches], [64, 64, 2])

    def test_snapshot_refresh_preserves_survivors_and_emits_only_changes(self):
        model, proxy = self._model()
        persistent = QtCore.QPersistentModelIndex(model.index(3, 0))
        resets, changed, inserted, removed = [], [], [], []
        model.modelReset.connect(lambda: resets.append(True))
        model.dataChanged.connect(lambda first, last, roles: changed.append((first.row(), last.row())))
        model.rowsInserted.connect(lambda parent, first, last: inserted.append(last - first + 1))
        model.rowsRemoved.connect(lambda parent, first, last: removed.append(last - first + 1))
        snapshot = [dict(rec) for rec in model._rows if rec['name'] != 'z.txt']
        snapshot[-1]['size'] = 99
        snapshot.append({'name': 'new.txt', 'path': 'C:/new.txt', 'is_dir': False, 'size': 1, 'mtime': 1})
        model.reconcile_snapshot(snapshot)
        proxy.sort(0, QtCore.Qt.AscendingOrder)
        self.assertEqual(resets, [])
        self.assertEqual(inserted, [1])
        self.assertEqual(removed, [1])
        self.assertEqual(len(changed), 1)
        self.assertEqual(persistent.data(), 'a.txt')
        self.assertEqual(model._rows[persistent.row()]['size'], 99)
        changed.clear()
        model.reconcile_snapshot([dict(rec) for rec in model._rows])
        self.assertEqual(changed, [])
        self.assertEqual(inserted, [1])
        self.assertEqual(removed, [1])


if __name__ == "__main__":
    unittest.main()
