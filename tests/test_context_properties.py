import unittest
from unittest import mock

import multipane_explorer as explorer


@unittest.skipUnless(explorer.HAS_PYWIN32, "Windows shell integration requires pywin32")
class ContextPropertiesTests(unittest.TestCase):
    def setUp(self):
        self.gui = self.enterContext(mock.patch.object(explorer, "win32gui"))
        self.gui.TrackPopupMenu.return_value = 13
        self.gui.GetMenuState.return_value = 0
        self.shell = self.enterContext(mock.patch.object(explorer, "shell"))
        self.enterContext(mock.patch.object(explorer, "_get_canonical_verb", return_value="properties"))
        self.enterContext(mock.patch.object(explorer, "_menu_item_text", return_value="Properties"))
        self.cm = mock.Mock()

    def invoke(self, paths):
        return explorer._invoke_menu(
            100, self.cm, 200, (20, 30), r"C:\selection",
            paths=paths, id_first=10, id_last=20,
        )

    def test_multiple_files_folders_and_mixed_selection_use_combined_context(self):
        for paths in (
            [r"C:\selection\first.txt", r"C:\selection\second.txt"],
            [r"C:\selection\folder1", r"C:\selection\folder2"],
            [r"C:\selection\first.txt", r"C:\elsewhere\folder"],
        ):
            with self.subTest(paths=paths):
                self.cm.reset_mock()
                self.assertTrue(self.invoke(paths))
                self.cm.InvokeCommand.assert_called_once_with(
                    (0, 100, 3, None, None, explorer.win32con.SW_SHOWNORMAL, 0, 0)
                )
                self.shell.SHObjectProperties.assert_not_called()
                self.shell.ShellExecuteEx.assert_not_called()

    def test_multi_selection_retries_canonical_verb_if_command_id_fails(self):
        self.cm.InvokeCommand.side_effect = [RuntimeError("command ID failed"), None]
        self.assertTrue(self.invoke([r"C:\first.txt", r"C:\second.txt"]))
        self.assertEqual(self.cm.InvokeCommand.call_count, 2)
        self.assertEqual(self.cm.InvokeCommand.call_args.args[0][2], "properties")
        self.shell.SHObjectProperties.assert_not_called()
        self.shell.ShellExecuteEx.assert_not_called()

    def test_multi_selection_failure_does_not_show_only_first_item(self):
        self.cm.InvokeCommand.side_effect = RuntimeError("shell unavailable")
        self.assertFalse(self.invoke([r"C:\first.txt", r"C:\second.txt"]))
        self.shell.SHObjectProperties.assert_not_called()
        self.shell.ShellExecuteEx.assert_not_called()

    def test_single_item_and_background_keep_single_path_properties(self):
        for paths, target in (([r"C:\selection\first.txt"], r"C:\selection\first.txt"),
                              (None, r"C:\selection")):
            with self.subTest(paths=paths):
                self.shell.reset_mock()
                self.assertTrue(self.invoke(paths))
                self.shell.SHObjectProperties.assert_called_once_with(100, 2, target, None)
                self.cm.InvokeCommand.assert_not_called()


if __name__ == "__main__":
    unittest.main()
