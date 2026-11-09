"""Per-number failures using simulated controls and isolated workbooks only."""
import unittest
from unittest.mock import Mock, patch

import open_qq_add_friend as adding
import qq_question_answer as answers
import qq_verification as verification
from task_control import StopRequested
from test_add_sequence import VirtualClock
from test_add_workflow import WorkflowHarness, fake_window
import test_excel_sheet_sequence as sheet_tests
from test_qq_verification import Desktop, Control, unconfigured_application, MESSAGE

NUMBER = '851984709'


class NumberFailureTests(unittest.TestCase):
    def ready(self, h):
        h.search_number.side_effect = lambda *args: ('results', 1)
        h.click_result_add_button.side_effect = lambda *args, **kwargs: True
        h.click_send_button.side_effect = lambda *args, **kwargs: True

    def test_every_ui_stage_failure_advances_to_next_number(self):
        for stage in ('open_add_friend_page', 'select_user_tab', 'search_number',
                      'click_result_add_button', 'prepare_verification_form',
                      'click_send_button', 'close_submitted_form'):
            with self.subTest(stage=stage), WorkflowHarness([fake_window(441)], target=441) as h:
                self.ready(h)
                action = getattr(h, stage)
                original = action.side_effect
                def fail_once(*args, **kwargs):
                    if action.call_count == 1:
                        raise RuntimeError('synthetic page refresh')
                    return original(*args, **kwargs) if callable(original) else action.return_value
                action.side_effect = fail_once
                h.run()
                self.assertEqual([c.args[2] for c in h.mark_rows_processed.call_args_list], [[1], [2], [3]])
                self.assertEqual(h.journal.clear.call_count, 3)
                sent = [c.kwargs['expected_number'] for c in h.click_send_button.call_args_list]
                self.assertEqual(len(sent), len(set(sent)))
                self.assertIn(h.numbers[2][0], sent)
                if stage != 'close_submitted_form':
                    reason = h.mark_rows_processed.call_args_list[0].kwargs['reason']
                    self.assertIn('synthetic page refresh', reason)
                    self.assertIn('待核对' if stage == 'click_send_button' else '未执行发送', reason)
                else:
                    reason = h.mark_rows_processed.call_args_list[0].kwargs['reason']
                    self.assertIn('申请已提交，关闭页面异常', reason)
                self.assertEqual(h.open_add_friend_page.call_count, 2)

    def test_cleanup_failure_does_not_prevent_next_search(self):
        with WorkflowHarness([fake_window(441)], target=441) as h:
            self.ready(h)
            h.prepare_verification_form.side_effect = [RuntimeError('multiple forms'), h.form, h.form]
            h.close_number_forms.side_effect = RuntimeError('stale cleanup')
            h.run()
            self.assertEqual(h.search_number.call_count, 3)
            self.assertEqual(h.click_send_button.call_count, 2)
            self.assertIn('stale cleanup', h.output.getvalue())

    def test_stop_and_mouse_failsafe_are_not_swallowed(self):
        class FailSafeException(Exception):
            pass
        for error in (StopRequested(), FailSafeException('stop')):
            for stage in ('prepare_verification_form', 'click_send_button', 'close_number_forms'):
                with self.subTest(error=type(error).__name__, stage=stage), WorkflowHarness([fake_window(441)], target=441) as h:
                    self.ready(h)
                    if stage == 'close_number_forms':
                        h.prepare_verification_form.side_effect = RuntimeError('unknown')
                    getattr(h, stage).side_effect = error
                    with self.assertRaises(type(error)):
                        h.run()
                    self.assertEqual(h.search_number.call_count, 1)

    def test_error_record_save_failure_stops_before_next_number(self):
        with WorkflowHarness([fake_window(441)], target=441) as h:
            self.ready(h)
            h.prepare_verification_form.side_effect = RuntimeError('unknown')
            h.mark_rows_processed.side_effect = PermissionError('disk locked')
            with self.assertRaisesRegex(RuntimeError, '本地处理记录'):
                h.run()
            h.search_number.assert_called_once()
            h.click_send_button.assert_not_called()
            h.journal.clear.assert_not_called()

    def test_later_numbers_use_confirmed_process_id_without_stale_wrapper_reads(self):
        with WorkflowHarness([fake_window(441)], target=441) as h:
            self.ready(h)
            def invalidate_wrapper(*args):
                h.windows[0].process_id.side_effect = RuntimeError('stale main wrapper')
                return 'results', 1
            h.search_number.side_effect = invalidate_wrapper
            h.run()
            self.assertEqual(h.click_send_button.call_count, 3)
            self.assertTrue(all(c.args[1] == 441 for c in h.click_send_button.call_args_list))


class FormReacquisitionTests(unittest.TestCase):
    def prepare(self, desktop, timeout=0):
        return answers.prepare_verification_form(desktop, 123, MESSAGE, timeout, expected_number=NUMBER)

    def test_other_numbers_and_accounts_are_not_counted_as_duplicates(self):
        target, *_ = unconfigured_application()
        old, *_ = unconfigured_application(number='123456789')
        other_account, *_ = unconfigured_application(pid=456)
        result = self.prepare(Desktop(old, other_account, target))
        self.assertIs(result.window, target)

    def test_duplicate_wrappers_of_same_native_window_are_deduplicated(self):
        target, *_ = unconfigured_application()
        other, *_ = unconfigured_application()
        target.handle = other.handle = 1234
        self.prepare(Desktop(target, other))

    def test_transient_empty_snapshot_waits_for_target(self):
        target, *_ = unconfigured_application()
        desktop = Mock()
        desktop.windows.side_effect = [[target], [], [target]]
        clock = VirtualClock()
        with patch.object(verification, 'time', clock):
            result = self.prepare(desktop, timeout=3)
        self.assertIs(result.window, target)
        self.assertGreaterEqual(clock.now, 2.0)

    def test_zero_forms_reports_missing_instead_of_multiple(self):
        target, *_ = unconfigured_application()
        desktop = Mock()
        desktop.windows.side_effect = [[target], []]
        with self.assertRaisesRegex(RuntimeError, '找到 0 个') as caught:
            self.prepare(desktop)
        self.assertNotIn('多个', str(caught.exception))

    def test_cleanup_closes_only_target_forms_without_sending(self):
        first, *_, send, cancel = unconfigured_application()
        second, *_ = unconfigured_application()
        old, *_ = unconfigured_application(number='123456789')
        other, *_ = unconfigured_application(pid=456)
        chat = Control('Window', 'QQ', children=[Control('Text', NUMBER)])
        for window in (first, second, old, other, chat):
            window.close = Mock(side_effect=lambda w=window: setattr(w, 'visible', False))
        verification.close_number_forms(Desktop(first, second, old, other, chat), 123, NUMBER)
        first.close.assert_called_once()
        second.close.assert_called_once()
        for window in (old, other, chat):
            window.close.assert_not_called()
        self.assertEqual(send.invoke_attempts, 0)
        self.assertEqual(cancel.invoke_attempts, 0)


class WorkbookFailureTests(unittest.TestCase):
    setUp = sheet_tests.SheetSequenceTests.setUp
    workbook = sheet_tests.SheetSequenceTests.workbook
    read = sheet_tests.SheetSequenceTests.read
    workflow = sheet_tests.SheetSequenceTests.workflow
    searched = sheet_tests.SheetSequenceTests.searched

    def test_unknown_form_error_marks_duplicates_and_continues_across_sheets(self):
        self.workbook([('First', [['100001'], ['100002']]),
                       ('Second', [['100001'], ['100003']])])
        with self.workflow() as h:
            h.search_number.side_effect = [('results', 1), ('no_results', 1), ('no_results', 1)]
            h.click_result_add_button.side_effect = lambda *args, **kwargs: True
            h.prepare_verification_form.side_effect = RuntimeError('multiple forms')
            h.run()
            self.assertEqual(self.searched(h), ['100001', '100002', '100003'])
            h.click_send_button.assert_not_called()
        book = self.read()[1]
        for title in ('First', 'Second'):
            self.assertEqual(book[title]['C1'].value, adding.PROCESSED_MARKER)
            self.assertIn('multiple forms', book[title]['D1'].value)
        with self.workflow() as h:
            h.run()
            h.search_number.assert_not_called()


if __name__ == '__main__':
    unittest.main()
