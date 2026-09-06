import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import multipane_explorer as e
import test_async_operations


class CustomizeTests(unittest.TestCase):
    run_gui_check = test_async_operations.AsyncOperationTests.run_gui_check

    def test_defaults_and_invalid_settings(self):
        for value in (None, {}, {'count': 3, 'items': [None]}):
            config = e.normalize_custom_commands(value)
            self.assertEqual(config['count'], 0)
            self.assertEqual(len(config['items']), 6)
        self.assertEqual(e.normalize_custom_commands({'items': [{'icon': 'AA'}]})['items'][0]['icon'], '')

    def test_layout_dialog_and_persistence(self):
        self.run_gui_check('''
            import tempfile
            from PyQt5 import QtCore, QtWidgets
            import multipane_explorer as e
            app = QtWidgets.QApplication([])
            with tempfile.TemporaryDirectory() as root:
                QtCore.QSettings.setDefaultFormat(QtCore.QSettings.IniFormat)
                for scope in (QtCore.QSettings.UserScope, QtCore.QSettings.SystemScope):
                    QtCore.QSettings.setPath(QtCore.QSettings.IniFormat, scope, root)
                window = e.MultiExplorer(4, [root] * 4)
                assert window.custom_commands['count'] == 0
                assert window.btn_customize.toolTip() == 'Customize command buttons'
                for count in (0, 2, 4, 6, 0):
                    window.custom_commands['count'] = count
                    for pane in window.panes:
                        pane._rebuild_custom_buttons()
                        grid = pane._tool_grid
                        assert grid.count() == 6 + count
                        assert grid.itemAtPosition(0, 0).widget() is pane.btn_cmd
                        assert grid.itemAtPosition(0, 1 + count // 2).widget() is pane.btn_explorer
                        assert grid.itemAtPosition(1, 1 + count // 2).widget() is pane.btn_new_file
                        for index, button in enumerate(pane._custom_buttons):
                            assert grid.itemAtPosition(index % 2, 1 + index // 2).widget() is button
                dialog = e.CustomizeDialog(window, window.custom_commands)
                assert all(not row[2].isEnabled() for row in dialog.editors)
                dialog.count_combo.setCurrentIndex(3)
                dialog.editors[5][1].setCurrentIndex(26)
                dialog.editors[5][2].setText('echo "kept"')
                assert all(row[2].isEnabled() for row in dialog.editors)
                dialog.count_combo.setCurrentIndex(0)
                config = dialog.config()
                assert config['items'][5] == {'icon': 'Z', 'command': 'echo "kept"'}
                settings = QtCore.QSettings(e.ORG_NAME, e.APP_NAME)
                settings.setValue('customize/commands', config); settings.sync()
                assert e.normalize_custom_commands(QtCore.QSettings(e.ORG_NAME, e.APP_NAME).value('customize/commands')) == config
                window._update_theme_dependent_icons()
                assert window.close()
                app.processEvents()
        ''')

    @unittest.skipUnless(os.name == 'nt', 'Windows CMD execution')
    def test_hidden_launch_flags(self):
        with tempfile.TemporaryDirectory() as root, mock.patch.object(e.subprocess, 'Popen') as launch:
            e.launch_custom_command('echo "hello world" > "out file.txt"', root)
            args, kwargs = launch.call_args
            self.assertIn('/d /s /c "echo "hello world" > "out file.txt""', args[0])
            self.assertEqual(kwargs['cwd'], root)
            self.assertTrue(kwargs['creationflags'] & subprocess.CREATE_NO_WINDOW)
            self.assertEqual(kwargs['startupinfo'].wShowWindow, subprocess.SW_HIDE)

    @unittest.skipUnless(os.name == 'nt', 'Windows CMD execution')
    def test_real_command_quoting_and_working_directory(self):
        with tempfile.TemporaryDirectory(prefix='custom command ') as root:
            script = Path(root, 'test script.cmd')
            script.write_text('@echo off\necho %~1>"output file.txt"\n')
            process = e.launch_custom_command('"test script.cmd" "hello world"', root)
            self.assertEqual(process.wait(timeout=5), 0)
            self.assertEqual(Path(root, 'output file.txt').read_text().strip(), 'hello world')
            process = e.launch_custom_command('echo first>first.txt && echo second>second.txt', root)
            self.assertEqual(process.wait(timeout=5), 0)
            self.assertTrue(Path(root, 'first.txt').exists())
            self.assertTrue(Path(root, 'second.txt').exists())


if __name__ == '__main__':
    unittest.main()
