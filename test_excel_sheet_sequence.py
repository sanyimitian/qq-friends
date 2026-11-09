"""Real temporary workbooks and mocked QQ boundaries; no live QQ actions."""
from contextlib import contextmanager
import io
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.teletype import extractText
from odf.text import P

import open_qq_add_friend as adding
from add_progress import AttemptJournal
from test_add_workflow import WorkflowHarness, fake_window
from task_control import StopRequested

READ = adding.read_number_rows
MARK = adding.mark_rows_processed


class SheetSequenceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='qq_sheet_sequence_')
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'numbers.xlsx'
        env = patch.dict(os.environ, QQ_ASSISTANT_STATE_DIR=str(Path(self.temp.name) / 'state'))
        env.start()
        self.addCleanup(env.stop)

    def workbook(self, sheets, active=0):
        book = Workbook()
        book.remove(book.active)
        for title, rows in sheets:
            sheet = book.create_sheet(title)
            for row in rows:
                sheet.append(row)
        book.active = active
        book.save(self.path)
        book.close()

    def read(self):
        numbers, book, sheet, path = adding.read_excel_number_rows(self.path)
        self.addCleanup(book.close)
        return numbers, book, sheet, book._qq_number_queue

    @contextmanager
    def workflow(self, quota=10, windows=None):
        windows = windows or [fake_window(441)]
        identities = ['111111111', '222222222'] * 2 if len(windows) > 1 else []
        with WorkflowHarness(windows, target=441 if len(windows) == 1 else 0,
                             per_account=quota, identities=identities) as harness:
            def read(*args):
                result = adding.read_excel_number_rows(self.path)
                self.addCleanup(result[1].close)
                harness.book, harness.queue = result[1], result[1]._qq_number_queue
                return result
            harness.read_number_rows.side_effect = read
            harness.mark_rows_processed.side_effect = MARK
            with (patch.object(adding, 'AttemptJournal', AttemptJournal),
                  patch.object(adding, 'NUMBER_FILE_PATH', self.path)):
                yield harness

    def searched(self, harness):
        return [item.args[1] for item in harness.search_number.call_args_list]

    def test_tag_order_ignores_active_and_skips_empty_and_group_headers(self):
        self.workbook([('First', [['group', '900001'], ['100001'], ['100002']]),
                       ('Empty', []), ('Title', [['group two', '900002']]),
                       ('Last', [['100003']])], active=3)
        original = self.path.read_bytes()
        numbers, book, sheet, queue = self.read()
        self.assertEqual([number for number, _ in numbers], ['100001', '100002', '100003'])
        self.assertEqual(sheet.title, 'First')
        self.assertEqual(book.active.title, 'Last')
        self.assertEqual(self.path.read_bytes(), original)

    def test_current_sheet_is_saved_before_next_search_and_tail_is_processed(self):
        self.workbook([('First', [['100001'], ['100002']]), ('Second', [['100003']])], active=1)
        with self.workflow(quota=10) as harness:
            def search(page, number, selected):
                if number == '100003':
                    check = load_workbook(self.path)
                    try:
                        self.assertEqual([check['First'].cell(r, 3).value for r in (1, 2)],
                                         [adding.PROCESSED_MARKER] * 2)
                    finally:
                        check.close()
                return 'no_results', 1
            harness.search_number.side_effect = search
            harness.run()
            self.assertEqual(self.searched(harness), ['100001', '100002', '100003'])
            self.assertEqual(harness.queue.pending_numbers(), [])
            self.assertIn('所有工作表', harness.output.getvalue())
        self.assertEqual(self.read()[1].active.title, 'Second')

    def test_quota_does_not_reset_on_sheet_transition(self):
        self.workbook([('First', [['100001'], ['100002']]),
                       ('Second', [['100003'], ['100004']])])
        with self.workflow(quota=3) as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100001', '100002', '100003'])
        self.assertEqual(self.read()[0], [('100004', [2])])
        with self.workflow(quota=3) as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100004'])

    def test_second_account_continues_unfinished_sheet_before_advancing(self):
        self.workbook([('First', [['100001'], ['100002'], ['100003']]),
                       ('Second', [['100004'], ['100005']])])
        with self.workflow(quota=2, windows=[fake_window(441), fake_window(442)]) as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100001', '100002', '100003', '100004'])
            self.assertEqual(harness.open_add_friend_page.call_count, 2)
        self.assertEqual(self.read()[0], [('100005', [2])])

    def test_cross_sheet_duplicate_sent_once_and_all_locations_written(self):
        self.workbook([('First', [['100001'], ['100002']]),
                       ('Second', [['100001'], ['100003'], ['100002']])])
        with self.workflow() as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100001', '100002', '100003'])
        _, book, _, queue = self.read()
        for sheet in book.worksheets:
            for row in range(1, sheet.max_row + 1):
                self.assertEqual(sheet.cell(row, 3).value, adding.PROCESSED_MARKER)

    def test_completed_duplicate_in_later_sheet_never_resends(self):
        self.workbook([('First', [['100001'], ['100002']]),
                       ('Second', [['100001', None, adding.PROCESSED_MARKER]])])
        with self.workflow() as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100002'])
        self.assertEqual(self.read()[1]['First']['C1'].value, adding.PROCESSED_MARKER)

    def test_timeout_is_skipped_before_next_sheet(self):
        self.workbook([('First', [['100001'], ['100002']]), ('Second', [['100003']])])
        with self.workflow() as harness:
            harness.search_number.side_effect = [('timeout', 3), ('no_results', 1), ('no_results', 1)]
            harness.run()
            self.assertEqual(self.searched(harness), ['100001', '100002', '100003'])
        _, book, _, _ = self.read()
        self.assertEqual(book['First']['C1'].value, adding.PROCESSED_MARKER)
        self.assertIn('超时', book['First']['D1'].value)
        self.assertEqual(book['Second']['C1'].value, adding.PROCESSED_MARKER)

    def test_unconfigured_form_is_skipped_and_next_sheet_runs(self):
        self.workbook([('First', [['100001']]), ('Second', [['100002']])])
        with self.workflow() as harness:
            harness.search_number.side_effect = lambda *args: ('results', 1)
            harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
            harness.prepare_verification_form.side_effect = adding.NoVerificationField(Mock())
            harness.run()
            self.assertEqual(self.searched(harness), ['100001', '100002'])
            harness.click_send_button.assert_not_called()

    def test_failed_save_stops_before_next_number_and_keeps_recovery_record(self):
        self.workbook([('First', [['100001']]), ('Second', [['100002']])])
        with self.workflow() as harness:
            harness.mark_rows_processed.side_effect = PermissionError('locked')
            with self.assertRaisesRegex(RuntimeError, '本地处理记录'):
                harness.run()
            self.assertEqual(self.searched(harness), ['100001'])
        journal = AttemptJournal(self.path, 'First')
        self.assertEqual(journal.records['100001']['state'], 'processed')
        with self.workflow() as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100002'])

    def test_recovery_uses_original_sheet_number_after_rows_and_active_change(self):
        self.workbook([('First', [['100001', None, adding.PROCESSED_MARKER]]),
                       ('Second', [['100003'], ['100002']]), ('Third', [['100002'], ['100004']])], active=0)
        journal = AttemptJournal(self.path, 'Second')
        journal.record('100002', [1], 'processed', reason='搜索无结果')
        with self.workflow() as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100003', '100004'])
        _, book, _, _ = self.read()
        self.assertEqual(book['Second']['C2'].value, adding.PROCESSED_MARKER)
        self.assertEqual(book['Third']['C1'].value, adding.PROCESSED_MARKER)
        self.assertFalse(journal.path.exists())



    def test_uncertain_sends_are_skipped_without_resending(self):
        for title in ['First', 'Second']:
            with self.subTest(sheet=title):
                first_status = adding.PROCESSED_MARKER if title == 'Second' else None
                self.workbook([('First', [['100001', None, first_status]]), ('Second', [['100002']])])
                number = '100001' if title == 'First' else '100002'
                journal = AttemptJournal(self.path, title)
                journal.record(number, [1], 'sending')
                with self.workflow() as harness:
                    harness.run()
                    self.assertEqual(self.searched(harness), ['100002'] if title == 'First' else [])
                    harness.click_send_button.assert_not_called()
                book = self.read()[1]
                self.assertEqual(book[title]['C1'].value, adding.PROCESSED_MARKER)
                self.assertIn('待核对', book[title]['D1'].value)
                self.assertEqual(AttemptJournal(self.path, title).records, {})

    def test_deferred_send_continues_same_sheet_and_later_sheet(self):
        self.workbook([('First', [['100001'], ['100002']]), ('Second', [['100003']])])
        journal = AttemptJournal(self.path, 'First')
        journal.record('100001', [1], 'sending')
        with self.workflow() as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100002', '100003'])
        book = self.read()[1]
        self.assertEqual(book['First']['C1'].value, adding.PROCESSED_MARKER)
        self.assertIn('待核对', book['First']['D1'].value)
        self.assertEqual(book['First']['C2'].value, adding.PROCESSED_MARKER)
        self.assertEqual(book['Second']['C1'].value, adding.PROCESSED_MARKER)

    def test_manually_confirmed_pending_record_releases_next_sheet(self):
        self.workbook([('First', [['100001', None, adding.PROCESSED_MARKER]]),
                       ('Second', [['100002']])])
        journal = AttemptJournal(self.path, 'First')
        journal.record('100001', [1], 'sending')
        with self.workflow() as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100002'])
        self.assertFalse(journal.path.exists())

    def test_new_uncertain_send_remains_durable_while_same_sheet_continues(self):
        self.workbook([('First', [['100001'], ['100002']]), ('Second', [['100003']])])
        with self.workflow() as harness:
            harness.search_number.side_effect = [('results', 1), ('no_results', 1), ('no_results', 1)]
            harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
            harness.click_send_button.side_effect = adding.SendUnconfirmed('待核对')
            harness.run()
            self.assertEqual(self.searched(harness), ['100001', '100002', '100003'])
            self.assertEqual(harness.click_send_button.call_count, 1)
        journal = AttemptJournal(self.path, 'First')
        self.assertEqual(journal.records, {})
        book = self.read()[1]
        self.assertEqual(book['First']['C1'].value, adding.PROCESSED_MARKER)
        self.assertIn('待核对', book['First']['D1'].value)
        self.assertEqual(book['First']['C2'].value, adding.PROCESSED_MARKER)
        self.assertEqual(book['Second']['C1'].value, adding.PROCESSED_MARKER)


    def test_deleted_recovery_sheet_is_not_replaced_by_same_number_elsewhere(self):
        self.workbook([('First', [['100001']]), ('Second', [['100002']])])
        journal = AttemptJournal(self.path, 'Deleted')
        journal.record('100002', [1], 'processed')
        with self.workflow() as harness:
            with self.assertRaisesRegex(RuntimeError, '恢复记录'):
                harness.run()
            harness.search_number.assert_not_called()

    def test_stop_mid_sheet_resumes_same_sheet_on_next_run(self):
        self.workbook([('First', [['100001'], ['100002']]), ('Second', [['100003']])])
        with self.workflow() as harness:
            harness.search_number.side_effect = [('no_results', 1), StopRequested()]
            with self.assertRaises(StopRequested):
                harness.run()
        with self.workflow() as harness:
            harness.run()
            self.assertEqual(self.searched(harness), ['100002', '100003'])

    def test_invalid_later_sheet_reports_name_and_row_without_writing(self):
        self.workbook([('First', [['100001']]), ('Bad', [['100002'], ['=100003']])])
        before = self.path.read_bytes()
        with self.assertRaisesRegex(ValueError, 'Bad.*2.*公式'):
            self.read()
        self.assertEqual(self.path.read_bytes(), before)

    def test_hidden_sheets_participate_in_tag_order(self):
        self.workbook([('First', [['100001']]), ('Hidden', [['100002']]), ('Last', [['100003']])])
        book = load_workbook(self.path)
        book['Hidden'].sheet_state = 'hidden'
        book.save(self.path)
        book.close()
        self.assertEqual([n for n, _ in self.read()[0]], ['100001', '100002', '100003'])


if __name__ == '__main__':
    unittest.main()
