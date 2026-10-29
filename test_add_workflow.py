"""Offline main-workflow regressions; never open QQ or touch user workbooks."""

import io
import unittest
from contextlib import ExitStack, redirect_stdout
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, call, patch

import open_qq_add_friend as adding


def fake_window(process):
    return SimpleNamespace(handle=process + 10000,
                           process_id=Mock(return_value=process))


class WorkflowHarness:
    """Replace every external boundary while exercising the real main()."""

    def __init__(self, windows, *, target=0, excluded=(), identities=(),
                 number_count=6, per_account=3, close_secondary=True):
        self.windows = windows
        self.events = []
        self.output = io.StringIO()
        self.identities = iter(identities)
        self.desktop = Mock(name='desktop')
        self.mouse = Mock(name='mouse')
        self.win32con = Mock(name='win32con')
        self.win32gui = Mock(name='win32gui')
        self.document = Mock(name='document')
        self.sheet = Mock(name='sheet')
        self.path = Path('synthetic_numbers.xlsx')
        self.journal = Mock(name='journal')
        self.journal.recover.side_effect = lambda numbers, callback, **kwargs: numbers
        self.form = Mock(name='verified_form')
        self.pages = {window.handle: Mock(name=f'page_{window.handle}')
                      for window in windows}
        self.numbers = [(str(123450000 + index), [index + 1])
                        for index in range(number_count)]
        self.dependencies = (Mock(return_value=self.desktop), self.mouse,
                             self.win32con, self.win32gui, *[Mock() for _ in range(6)])
        self.values = {
            'TARGET_PROCESS_ID': target,
            'EXCLUDED_QQ_ACCOUNTS': set(excluded),
            'NUMBERS_PER_ACCOUNT': per_account,
            'CLOSE_SECONDARY_WINDOWS': close_secondary,
            'AUTO_SEND': True,
        }

    def _read_identity(self, window, *args):
        self.events.append(('identity', window.handle))
        value = next(self.identities)
        if isinstance(value, BaseException):
            raise value
        return value

    def _read_numbers(self, *args):
        self.events.append(('numbers',))
        return self.numbers, self.document, self.sheet, self.path

    def _activate(self, window, *args):
        self.events.append(('activate', window.handle))

    def _cleanup(self, desktop, windows, *args):
        self.events.append(('cleanup', tuple(window.handle for window in windows)))

    def _open_page(self, desktop, con, gui, window=None):
        self.events.append(('open', window.handle))
        return window, self.pages[window.handle]

    def _search(self, page, number, selected):
        self.events.append(('search', number))
        return 'no_results', 1

    def _mark(self, document, sheet, rows, *args):
        self.events.append(('mark', tuple(rows)))

    def __enter__(self):
        self.stack = ExitStack()
        self.stack.enter_context(patch.object(adding.sys, 'platform', 'win32'))
        self.stack.enter_context(patch.multiple(adding, **self.values))
        replacements = {
            'ensure_number_file_writable': Mock(),
            'AttemptJournal': Mock(return_value=self.journal),
            'current_sheet_numbers': Mock(return_value={number for number, _ in self.numbers}),
            'prepare_verification_form': Mock(return_value=self.form),
            'cancel_unconfigured_form': Mock(return_value=True),
            'load_dependencies': Mock(return_value=self.dependencies),
            'find_qq_windows': Mock(return_value=self.windows),
            'read_qq_account_from_avatar': Mock(side_effect=self._read_identity),
            'read_number_rows': Mock(side_effect=self._read_numbers),
            'activate_qq_account_window': Mock(side_effect=self._activate),
            'close_qq_secondary_windows': Mock(side_effect=self._cleanup),
            'open_add_friend_page': Mock(side_effect=self._open_page),
            'select_user_tab': Mock(return_value=True),
            'search_number': Mock(side_effect=self._search),
            'mark_rows_processed': Mock(side_effect=self._mark),
            'click_result_add_button': Mock(side_effect=AssertionError(
                'This harness must never click an add button.')),
            'click_send_button': Mock(side_effect=AssertionError(
                'This harness must never send a request.')),
        }
        for name, replacement in replacements.items():
            setattr(self, name, replacement)
            self.stack.enter_context(patch.object(adding, name, replacement))
        self.stack.enter_context(patch.object(adding.time, 'sleep'))
        self.stack.enter_context(redirect_stdout(self.output))
        return self

    def __exit__(self, *args):
        return self.stack.__exit__(*args)

    def run(self):
        adding.main()


class MainWorkflowTests(unittest.TestCase):
    def assert_no_search(self, harness):
        harness.open_add_friend_page.assert_not_called()
        harness.search_number.assert_not_called()
        harness.mark_rows_processed.assert_not_called()
        harness.click_result_add_button.assert_not_called()
        harness.click_send_button.assert_not_called()

    def test_selected_window_completes_batch_without_reading_profile(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441) as harness:
            harness.read_qq_account_from_avatar.side_effect = AssertionError(
                'An explicitly selected single window needs no profile read.')
            harness.run()
            harness.read_qq_account_from_avatar.assert_not_called()
            harness.open_add_friend_page.assert_called_once_with(
                harness.desktop, harness.win32con, harness.win32gui, window)
            self.assertEqual(harness.search_number.call_args_list, [
                call(harness.pages[window.handle], number, True)
                for number, _ in harness.numbers[:3]])
            self.assertEqual(harness.mark_rows_processed.call_count, 3)
            self.assertEqual([event[1] for event in harness.events if event[0] == 'mark'],
                             [(1,), (2,), (3,)])
            self.assertLess(harness.events.index(('activate', window.handle)),
                            harness.events.index(('open', window.handle)))
            harness.close_qq_secondary_windows.assert_called_once_with(
                harness.desktop, [window], harness.win32con, harness.win32gui)

    def test_selected_window_with_exclusions_reads_identity_once_before_cleanup(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441, excluded={'999999999'},
                             identities=['111111111']) as harness:
            harness.run()
            harness.read_qq_account_from_avatar.assert_called_once()
            self.assertEqual(harness.search_number.call_count, 3)
            self.assertLess(harness.events.index(('identity', window.handle)),
                            harness.events.index(('cleanup', (window.handle,))))

    def test_selected_excluded_account_never_opens_search_or_loads_numbers(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441, excluded={'111111111'},
                             identities=['111111111']) as harness:
            harness.run()
            harness.read_qq_account_from_avatar.assert_called_once()
            harness.read_number_rows.assert_not_called()
            harness.activate_qq_account_window.assert_not_called()
            harness.close_qq_secondary_windows.assert_not_called()
            self.assert_no_search(harness)

    def test_auto_detected_single_window_reads_identity_once(self):
        window = fake_window(441)
        with WorkflowHarness([window], identities=['111111111']) as harness:
            harness.run()
            harness.read_qq_account_from_avatar.assert_called_once()
            self.assertEqual(harness.search_number.call_count, 3)
            self.assertEqual(harness.mark_rows_processed.call_count, 3)

    def test_multi_account_mismatch_stops_before_cleanup_or_search(self):
        first, second = fake_window(441), fake_window(442)
        with WorkflowHarness([first, second],
                             identities=['111111111', '222222222', '999999999']) as harness:
            with self.assertRaises(RuntimeError):
                harness.run()
            self.assertEqual(harness.read_qq_account_from_avatar.call_count, 3)
            harness.activate_qq_account_window.assert_called_once()
            harness.close_qq_secondary_windows.assert_not_called()
            self.assert_no_search(harness)

    def test_multi_account_verifies_before_closing_current_account_secondary_windows(self):
        first, second = fake_window(441), fake_window(442)
        with WorkflowHarness([first, second],
                             identities=['111111111', '222222222',
                                         '111111111', '222222222']) as harness:
            harness.run()
            self.assertEqual(harness.read_qq_account_from_avatar.call_count, 4)
            self.assertEqual(harness.search_number.call_count, 6)
            self.assertEqual(harness.mark_rows_processed.call_count, 6)
            self.assertEqual(harness.close_qq_secondary_windows.call_args_list, [
                call(harness.desktop, [window], harness.win32con, harness.win32gui)
                for window in (first, second)])
            for window in (first, second):
                activate_index = harness.events.index(('activate', window.handle))
                verify_index = harness.events.index(('identity', window.handle), activate_index)
                cleanup_index = harness.events.index(('cleanup', (window.handle,)))
                open_index = harness.events.index(('open', window.handle))
                self.assertLess(activate_index, verify_index)
                self.assertLess(verify_index, cleanup_index)
                self.assertLess(cleanup_index, open_index)

    def test_multi_account_exclusions_preserve_allowed_window_order(self):
        first, excluded, third = (fake_window(value) for value in (441, 442, 443))
        with WorkflowHarness([first, excluded, third], excluded={'222222222'},
                             identities=['111111111', '222222222', '333333333',
                                         '111111111', '333333333']) as harness:
            harness.run()
            self.assertEqual([args.args[-1] for args in harness.open_add_friend_page.call_args_list],
                             [first, third])
            self.assertEqual(harness.mark_rows_processed.call_count, 6)
            self.assertNotIn(('activate', excluded.handle), harness.events)
            self.assertNotIn(('cleanup', (excluded.handle,)), harness.events)

    def test_duplicate_account_window_does_not_consume_second_batch(self):
        first, duplicate, third = (fake_window(value) for value in (441, 442, 443))
        with WorkflowHarness([first, duplicate, third],
                             identities=['111111111', '111111111', '333333333',
                                         '111111111', '333333333']) as harness:
            harness.run()
            self.assertEqual([args.args[-1] for args in harness.open_add_friend_page.call_args_list],
                             [first, third])
            self.assertEqual(harness.mark_rows_processed.call_count, 6)
            self.assertNotIn(('activate', duplicate.handle), harness.events)

    def test_selected_window_empty_numbers_does_not_touch_ui(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441, number_count=0) as harness:
            harness.run()
            harness.read_qq_account_from_avatar.assert_not_called()
            harness.activate_qq_account_window.assert_not_called()
            harness.close_qq_secondary_windows.assert_not_called()
            self.assert_no_search(harness)

    def test_selected_window_insufficient_numbers_does_not_touch_ui(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441, number_count=2) as harness:
            harness.run()
            harness.read_qq_account_from_avatar.assert_not_called()
            harness.activate_qq_account_window.assert_not_called()
            harness.close_qq_secondary_windows.assert_not_called()
            self.assert_no_search(harness)

    def test_required_identity_failure_does_not_bypass_exclusion_checks(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441, excluded={'111111111'},
                             identities=[RuntimeError('synthetic profile failure')]) as harness:
            with self.assertRaises(RuntimeError):
                harness.run()
            harness.read_number_rows.assert_not_called()
            harness.activate_qq_account_window.assert_not_called()
            harness.close_qq_secondary_windows.assert_not_called()
            self.assert_no_search(harness)

    def test_failed_activation_stops_before_cleanup_or_search(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441) as harness:
            harness.activate_qq_account_window.side_effect = RuntimeError('not foreground')
            with self.assertRaises(RuntimeError):
                harness.run()
            harness.close_qq_secondary_windows.assert_not_called()
            self.assert_no_search(harness)

    def test_cleanup_setting_can_be_disabled_for_selected_window(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441, close_secondary=False) as harness:
            harness.run()
            harness.close_qq_secondary_windows.assert_not_called()
            harness.read_qq_account_from_avatar.assert_not_called()
            self.assertEqual(harness.search_number.call_count, 3)

    def test_auto_send_disabled_leaves_open_application_unmarked(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441) as harness, \
             patch.object(adding, 'AUTO_SEND', False):
            harness.search_number.side_effect = None
            harness.search_number.return_value = ('results', 1)
            harness.click_result_add_button.side_effect = None
            harness.click_result_add_button.return_value = True
            harness.run()
            harness.search_number.assert_called_once()
            harness.click_result_add_button.assert_called_once_with(
                harness.pages[window.handle], expected_number=harness.numbers[0][0])
            harness.click_send_button.assert_not_called()
            harness.mark_rows_processed.assert_not_called()
            harness.prepare_verification_form.assert_called_once()
            harness.journal.record.assert_not_called()

    def test_selected_batch_marks_only_after_mocked_send(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441) as harness:
            harness.search_number.side_effect = lambda page, number, selected: ('results', 1)
            harness.click_result_add_button.side_effect = lambda page, **kwargs: True
            def sent(*args, **kwargs):
                harness.events.append(('send',))
                return True
            harness.click_send_button.side_effect = sent
            harness.run()
            self.assertEqual(harness.search_number.call_count, 3)
            self.assertEqual(harness.click_result_add_button.call_count, 3)
            self.assertEqual(harness.click_send_button.call_count, 3)
            self.assertEqual([event for event in harness.events if event[0] in {'send', 'mark'}],
                             [('send',), ('mark', (1,)), ('send',), ('mark', (2,)),
                              ('send',), ('mark', (3,))])
            self.assertEqual(harness.prepare_verification_form.call_count, 3)
            self.assertEqual(harness.journal.record.call_args_list, [
                item for number, rows in harness.numbers[:3]
                for item in (call(number, rows, 'sending'), call(number, rows, 'processed'))])
            self.assertEqual(harness.journal.clear.call_count, 3)

    def test_locked_workbook_stops_before_any_qq_access(self):
        with WorkflowHarness([fake_window(441)], target=441) as harness:
            harness.ensure_number_file_writable.side_effect = PermissionError('close Excel')
            with self.assertRaises(PermissionError):
                harness.run()
            harness.load_dependencies.assert_not_called()
            harness.find_qq_windows.assert_not_called()
            harness.read_qq_account_from_avatar.assert_not_called()
            self.assert_no_search(harness)

    def test_verification_failure_never_sends_or_marks(self):
        with WorkflowHarness([fake_window(441)], target=441) as harness:
            harness.search_number.side_effect = lambda *args: ('results', 1)
            harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
            harness.prepare_verification_form.side_effect = RuntimeError('message truncated')
            with self.assertRaisesRegex(RuntimeError, 'truncated'):
                harness.run()
            harness.click_send_button.assert_not_called()
            harness.mark_rows_processed.assert_not_called()
            harness.journal.record.assert_not_called()

    def test_missing_verification_field_cancels_only_that_number_and_continues(self):
        window = fake_window(441)
        with WorkflowHarness([window], target=441) as harness:
            form = Mock(name='same_number_unconfigured_form')
            harness.search_number.side_effect = lambda *args: ('results', 1)
            harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
            harness.prepare_verification_form.side_effect = [
                adding.NoVerificationField(form), harness.form, harness.form,
            ]
            harness.click_send_button.side_effect = lambda *args, **kwargs: True
            harness.run()
            harness.cancel_unconfigured_form.assert_called_once_with(
                form, window.process_id(), harness.numbers[0][0], adding.SEND_TIMEOUT_SECONDS)
            self.assertEqual(harness.search_number.call_count, 3)
            harness.click_send_button.assert_has_calls([
                call(harness.desktop, window.process_id(), adding.SEND_TIMEOUT_SECONDS,
                     form=harness.form, expected_number=number)
                for number, _ in harness.numbers[1:3]])
            self.assertEqual([item.args[2] for item in harness.mark_rows_processed.call_args_list],
                             [[2], [3]])
            self.assertEqual(harness.journal.record.call_args_list, [
                item for number, rows in harness.numbers[1:3]
                for item in (call(number, rows, 'sending'), call(number, rows, 'processed'))])
            self.assertNotIn(harness.numbers[0][0],
                             [item.args[0] for item in harness.journal.record.call_args_list])

    def test_unconfigured_form_cancel_failure_stops_without_marking(self):
        with WorkflowHarness([fake_window(441)], target=441) as harness:
            form = Mock(name='unconfigured_form')
            harness.search_number.side_effect = lambda *args: ('results', 1)
            harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
            harness.prepare_verification_form.side_effect = adding.NoVerificationField(form)
            harness.cancel_unconfigured_form.side_effect = RuntimeError('cancel unconfirmed')
            with self.assertRaisesRegex(RuntimeError, 'cancel unconfirmed'):
                harness.run()
            harness.click_send_button.assert_not_called()
            harness.mark_rows_processed.assert_not_called()
            harness.journal.record.assert_not_called()

    def test_journal_write_failure_prevents_send(self):
        with WorkflowHarness([fake_window(441)], target=441) as harness:
            harness.search_number.side_effect = lambda *args: ('results', 1)
            harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
            harness.journal.record.side_effect = OSError('disk full')
            with self.assertRaises(OSError):
                harness.run()
            harness.click_send_button.assert_not_called()
            harness.mark_rows_processed.assert_not_called()

    def test_save_failure_after_send_keeps_recovery_record_and_stops_batch(self):
        with WorkflowHarness([fake_window(441)], target=441) as harness:
            harness.search_number.side_effect = lambda *args: ('results', 1)
            harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
            harness.click_send_button.side_effect = lambda *args, **kwargs: True
            harness.mark_rows_processed.side_effect = PermissionError('locked')
            with self.assertRaisesRegex(RuntimeError, '已保留.*本地处理记录'):
                harness.run()
            harness.click_send_button.assert_called_once()
            harness.search_number.assert_called_once()
            harness.journal.record.assert_has_calls([
                call(harness.numbers[0][0], [1], 'sending'),
                call(harness.numbers[0][0], [1], 'processed'),
            ])
            harness.journal.clear.assert_not_called()

    def test_stop_during_send_retains_uncertain_record_without_marking(self):
        from task_control import StopRequested
        with WorkflowHarness([fake_window(441)], target=441) as harness:
            harness.search_number.side_effect = lambda *args: ('results', 1)
            harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
            harness.click_send_button.side_effect = StopRequested()
            with self.assertRaises(StopRequested):
                harness.run()
            harness.journal.record.assert_called_once_with(harness.numbers[0][0], [1], 'sending')
            harness.journal.clear.assert_not_called()
            harness.mark_rows_processed.assert_not_called()

    def test_recovered_numbers_are_removed_before_searching(self):
        with WorkflowHarness([fake_window(441)], target=441) as harness:
            harness.journal.recover.side_effect = lambda numbers, callback, **kwargs: numbers[1:]
            harness.run()
            self.assertEqual([args.args[1] for args in harness.search_number.call_args_list],
                             [number for number, _ in harness.numbers[1:4]])

    def test_missing_selected_window_stops_before_profile_or_number_loading(self):
        with WorkflowHarness([], target=441) as harness:
            with self.assertRaises(RuntimeError) as caught:
                harness.run()
            self.assertIn('441', str(caught.exception))
            harness.read_qq_account_from_avatar.assert_not_called()
            harness.read_number_rows.assert_not_called()
            harness.activate_qq_account_window.assert_not_called()
            harness.close_qq_secondary_windows.assert_not_called()
            self.assert_no_search(harness)

    def test_missing_auto_detected_window_stops_before_profile_or_number_loading(self):
        with WorkflowHarness([]) as harness:
            with self.assertRaises(RuntimeError):
                harness.run()
            harness.read_qq_account_from_avatar.assert_not_called()
            harness.read_number_rows.assert_not_called()
            harness.activate_qq_account_window.assert_not_called()
            harness.close_qq_secondary_windows.assert_not_called()
            self.assert_no_search(harness)


if __name__ == '__main__':
    unittest.main()
