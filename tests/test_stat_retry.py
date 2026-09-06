import unittest
from unittest import mock

from PyQt5 import QtCore
import multipane_explorer as explorer


class StatRetryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtCore.QCoreApplication.instance() or QtCore.QCoreApplication([])

    def models(self):
        for model in (explorer.FastDirModel(), explorer.SearchResultModel()):
            model.append_rows([{"name": "gone.txt", "path": "gone.txt", "is_dir": False}])
            yield model

    def test_failed_stats_back_off_and_stop_after_three_attempts(self):
        for model in self.models():
            with self.subTest(model=type(model).__name__), mock.patch.object(explorer.time, "monotonic", return_value=100) as clock:
                rec = model._rows[0]
                for attempt in range(3):
                    self.assertEqual(explorer._stat_retry_delay_ms(rec), 0)
                    model.apply_stat_batch([("gone.txt", None, None)])
                    self.assertFalse(model.has_stat(0))
                    if attempt < 2:
                        self.assertGreater(explorer._stat_retry_delay_ms(rec), 0)
                    clock.return_value += 10
                self.assertIsNone(explorer._stat_retry_delay_ms(rec))
                self.assertIsNone(rec.get("size"))

    def test_successful_retry_clears_failure_and_restores_metadata(self):
        for model in self.models():
            model.apply_stat_batch([("gone.txt", None, None)])
            model.apply_stat_batch([("gone.txt", 42, 1000)])
            self.assertTrue(model.has_stat(0))
            self.assertNotIn("stat_failures", model._rows[0])
            self.assertEqual(model._rows[0]["size"], 42)


if __name__ == "__main__":
    unittest.main()
