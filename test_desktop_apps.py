import json
import os
from pathlib import Path
import re
import tempfile
import time
import unittest
from unittest.mock import patch

import gui_config
from gui_controller import TaskController
import zq_qq_001 as capture_script


class DesktopTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='qq_desktop_test_')
        self.env = patch.dict(os.environ, {'QQ_ASSISTANT_STATE_DIR': self.temp.name})
        self.env.start()

    def tearDown(self):
        self.env.stop()
        self.temp.cleanup()

    def test_defaults_and_invalid_values(self):
        for mode in ('capture', 'add'):
            config = gui_config.validate(mode, gui_config.defaults(mode))
            gui_config.save_config(mode, config)
            self.assertEqual(config, gui_config.load_config(mode))
        capture = gui_config.defaults('capture')
        self.assertEqual(capture['TARGET_QQ_NUMBER'], '3805745874')
        self.assertEqual(capture['MAX_GROUPS'], 2)
        self.assertEqual(capture['MAX_MEMBERS_PER_GROUP'], 2)
        for invalid_number in ('', '1234', '1234567890123', '12a456789'):
            invalid_config = dict(capture, TARGET_QQ_NUMBER=invalid_number)
            with self.subTest(target_qq_number=invalid_number):
                with self.assertRaises(ValueError):
                    gui_config.validate('capture', invalid_config)
        self.assertTrue(re.search(capture['GROUP_HEADER_PATTERN'], '测试群 (100)'))
        capture['MAX_GROUPS'] = '-1'
        with self.assertRaises(ValueError):
            gui_config.validate('capture', capture)
        adding = gui_config.defaults('add')
        adding['NUMBER_INTERVAL_SECONDS'] = 'nan'
        with self.assertRaises(ValueError):
            gui_config.validate('add', adding)

    def test_capture_selects_only_the_configured_qq_number(self):
        first = type('Window', (), {'process_id': lambda self: 101})()
        target = type('Window', (), {'process_id': lambda self: 202})()
        config = gui_config.validate(
            'capture', dict(gui_config.defaults('capture'), TARGET_QQ_NUMBER='987654321')
        )
        original_target = capture_script.TARGET_QQ_NUMBER
        self.addCleanup(setattr, capture_script, 'TARGET_QQ_NUMBER', original_target)
        gui_config.apply_config('capture', capture_script, config)
        accounts = {101: '123456789', 202: '987654321'}
        with patch.object(capture_script, 'find_qq_windows', return_value=[first, target]), \
                patch.object(capture_script, 'read_qq_account_from_avatar',
                             side_effect=lambda window, *_: accounts[window.process_id()]):
            selected = capture_script.select_qq_window_by_number(None, None, None, None)
        self.assertIs(selected, target)

    def test_capture_stops_when_configured_qq_number_is_not_logged_in(self):
        window = type('Window', (), {'process_id': lambda self: 101})()
        config = gui_config.validate(
            'capture', dict(gui_config.defaults('capture'), TARGET_QQ_NUMBER='987654321')
        )
        original_target = capture_script.TARGET_QQ_NUMBER
        self.addCleanup(setattr, capture_script, 'TARGET_QQ_NUMBER', original_target)
        gui_config.apply_config('capture', capture_script, config)
        with patch.object(capture_script, 'find_qq_windows', return_value=[window]), \
                patch.object(capture_script, 'read_qq_account_from_avatar', return_value='123456789'):
            with self.assertRaisesRegex(RuntimeError, '没有找到登录 QQ 号为 987654321'):
                capture_script.select_qq_window_by_number(None, None, None, None)

    def test_ods_repeated_rows_and_atomic_save(self):
        from odf.opendocument import OpenDocumentSpreadsheet, load
        from odf.table import Table, TableRow, TableCell
        from odf.text import P
        from odf.teletype import extractText
        import open_qq_add_friend as adding
        document = OpenDocumentSpreadsheet()
        sheet = Table(name='QQ号码')
        document.spreadsheet.addElement(sheet)
        sheet.addElement(TableRow(numberrowsrepeated=3))
        row = TableRow(numberrowsrepeated=2)
        row.addElement(TableCell(valuetype='string', stringvalue='123456789'))
        sheet.addElement(row)
        path = Path(self.temp.name) / 'numbers.ods'
        document.save(str(path))
        with patch.object(adding, 'NUMBER_FILE_PATH', path):
            numbers, doc, first, _ = adding.read_number_rows(load, Table, TableRow, TableCell, extractText)
            self.assertEqual(numbers, [('123456789', [4, 5])])
            adding.mark_rows_processed(doc, first, [4, 5], TableRow, TableCell, P, path)
            self.assertEqual(adding.read_number_rows(load, Table, TableRow, TableCell, extractText)[0], [])
        self.assertEqual(list(path.parent.glob('.qq_save_*')), [])

    def test_ods_ignores_column_b_and_writes_status_to_column_c(self):
        from odf.opendocument import OpenDocumentSpreadsheet, load
        from odf.table import Table, TableRow, TableCell
        from odf.text import P
        from odf.teletype import extractText
        import open_qq_add_friend as adding
        document = OpenDocumentSpreadsheet()
        sheet = Table(name='QQ号码')
        document.spreadsheet.addElement(sheet)
        row = TableRow()
        row.addElement(TableCell(valuetype='string', stringvalue='123456789'))
        row.addElement(TableCell(valuetype='string', stringvalue='群名称+群号'))
        row.addElement(TableCell())
        sheet.addElement(row)
        path = Path(self.temp.name) / 'numbers_columns.ods'
        document.save(str(path))
        with patch.object(adding, 'NUMBER_FILE_PATH', path):
            numbers, doc, active, _ = adding.read_number_rows(load, Table, TableRow, TableCell, extractText)
            self.assertEqual(numbers, [('123456789', [1])])
            adding.mark_rows_processed(
                doc, active, [1], TableRow, TableCell, P, path,
                reason='搜索结果中没有可点击的添加按钮',
            )
            check = load(str(path))
            check_row = check.spreadsheet.getElementsByType(Table)[0].getElementsByType(TableRow)[0]
            cells = check_row.getElementsByType(TableCell)
            self.assertEqual(adding.normalize_cell_value(cells[1], extractText), '群名称+群号')
            self.assertEqual(adding.normalize_cell_value(cells[2], extractText), adding.PROCESSED_MARKER)
            self.assertEqual(
                adding.normalize_cell_value(cells[3], extractText),
                '加好友跳过原因：搜索结果中没有可点击的添加按钮',
            )
            self.assertEqual(adding.read_number_rows(load, Table, TableRow, TableCell, extractText)[0], [])

    def await_result(self, controller, timeout=12):
        end = time.monotonic() + timeout
        output = ''
        while time.monotonic() < end:
            text, result = controller.poll()
            output += text
            if result:
                return output, result
            time.sleep(0.05)
        self.fail('Worker did not finish')

    def await_started(self, controller, timeout=12):
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            text, result = controller.poll()
            if '模拟任务已启动' in text:
                return
            if result:
                self.fail(result)
            time.sleep(0.05)
        self.fail('Worker did not start')

    def test_stop_restart_and_exclusive_task(self):
        first = TaskController('capture')
        second = TaskController('add')
        try:
            first.start(gui_config.defaults('capture'), simulate=True)
            self.await_started(first)
            with self.assertRaises(RuntimeError):
                first.start(gui_config.defaults('capture'), simulate=True)
            second.start(gui_config.defaults('add'), simulate=True)
            _, result = self.await_result(second)
            self.assertEqual(result['status'], 'error')
            self.assertIn('另一个 QQ 任务', result['message'])
            started = time.monotonic()
            first.stop()
            _, result = self.await_result(first)
            self.assertEqual(result['status'], 'stopped')
            self.assertLess(time.monotonic() - started, 3)
            first.start(gui_config.defaults('capture'), simulate=True)
            self.await_started(first)
            first.stop()
            _, result = self.await_result(first)
            self.assertEqual(result['status'], 'stopped')
            second.start(gui_config.defaults('add'), simulate=True)
            self.await_started(second)
            second.stop()
            _, result = self.await_result(second)
            self.assertEqual(result['status'], 'stopped')
        finally:
            first.cleanup()
            second.cleanup()

    def test_forced_stop_and_close_cleanup(self):
        controller = TaskController('capture')
        try:
            controller.start(gui_config.defaults('capture'), simulate=True)
            self.await_started(controller)
            controller.stop()
            controller.force_stop()
            _, result = self.await_result(controller)
            self.assertEqual(result['status'], 'stopped')
            controller.start(gui_config.defaults('capture'), simulate=True)
            self.await_started(controller)
            process = controller.process
            controller.cleanup()
            self.assertIsNotNone(process.poll())
        finally:
            controller.cleanup()


if __name__ == '__main__':
    unittest.main(verbosity=2)
