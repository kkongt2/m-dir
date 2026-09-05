import unittest

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


if __name__ == "__main__":
    unittest.main()
