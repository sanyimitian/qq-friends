"""Send completion and ambiguous-send recovery; synthetic controls only."""
from contextlib import ExitStack, redirect_stdout
import io
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import qq_verification as v
import qq_question_answer as q
import open_qq_add_friend as adding
from task_control import StopRequested
from test_qq_verification import application, unconfigured_application, Control, MESSAGE, Desktop
from test_add_sequence import VirtualClock
from test_add_workflow import WorkflowHarness, fake_window

NUMBER = '851984709'


class SendResilienceTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.stack.enter_context(redirect_stdout(io.StringIO()))
        self.clock = VirtualClock()
        self.stack.enter_context(patch.object(v, 'time', self.clock))
        self.window, self.edit, _, _, self.send = application()
        self.edit.value = MESSAGE

    def submit(self, timeout=5):
        return v.click_verified_send(self.window, 123, MESSAGE, expected_number=NUMBER, timeout=timeout)

    def keep_open(self, target):
        target.pending_close = None

    def show_after_send(self, text):
        def show(target):
            self.keep_open(target)
            target._children.append(Control('Text', text))
        self.send.on_invoke = show

    def test_slow_close_after_fifteen_seconds_succeeds_once(self):
        self.send.on_invoke = self.keep_open
        self.clock.on_sleep = lambda clock: setattr(self.window, 'visible', clock.now < 15)
        self.assertTrue(self.submit())
        self.assertGreaterEqual(self.clock.now, 15)
        self.assertEqual(self.send.invoke_attempts, 1)

    def test_success_message_is_accepted_even_while_window_remains_open(self):
        self.show_after_send('申请已发送')
        self.assertTrue(self.submit())
        self.assertTrue(self.window.visible)
        self.assertEqual(self.send.invoke_attempts, 1)

    def test_existing_success_text_is_not_new_confirmation(self):
        self.window._children.append(Control('Text', '申请已发送'))
        self.send.on_invoke = self.keep_open
        with self.assertRaises(v.SendUnconfirmed):
            self.submit(timeout=0)

    def test_success_in_edit_value_is_not_confirmation(self):
        self.send.on_invoke = self.keep_open
        self.window._children.append(Control('Edit', '申请已发送'))
        with self.assertRaises(v.SendUnconfirmed):
            self.submit(timeout=0)

    def test_explicit_failure_is_not_success_or_resubmitted(self):
        self.show_after_send('发送失败')
        with self.assertRaises(v.SendRejected) as caught:
            self.submit()
        self.assertFalse(caught.exception.halt_run)
        self.assertEqual(self.send.invoke_attempts, 1)

    def test_frequency_restriction_requests_a_stop(self):
        self.show_after_send('操作过于频繁')
        with self.assertRaises(v.SendRejected) as caught:
            self.submit()
        self.assertTrue(caught.exception.halt_run)

    def test_existing_failure_blocks_first_send(self):
        self.window._children.append(Control('Text', '发送失败'))
        with self.assertRaisesRegex(RuntimeError, '未点击发送'):
            self.submit()
        self.assertEqual(self.send.invoke_attempts, 0)

    def test_busy_form_gets_one_bounded_extension(self):
        self.show_after_send('正在提交')
        with self.assertRaises(v.SendUnconfirmed):
            self.submit()
        self.assertGreaterEqual(self.clock.now, 30)
        self.assertLess(self.clock.now, 31)
        self.assertEqual(self.send.invoke_attempts, 1)

    def test_transient_visibility_failure_is_retried_then_closes(self):
        def after(target):
            self.keep_open(target)
            target.visible_error = RuntimeError('transient')
        self.send.on_invoke = after
        def update(clock):
            if clock.now >= 1:
                self.window.visible_error = None
                self.window.visible = False
        self.clock.on_sleep = update
        self.assertTrue(self.submit())
        self.assertEqual(self.send.invoke_attempts, 1)

    def test_invoke_error_after_actual_close_is_confirmed_without_repeat(self):
        def invoked():
            self.window.visible = False
            raise RuntimeError('completion raced')
        action = Mock(Invoke=Mock(side_effect=invoked))
        with patch.object(Control, 'iface_invoke', new=property(lambda control: action)):
            self.assertTrue(self.submit())
        action.Invoke.assert_called_once()

    def test_stop_during_confirmation_propagates(self):
        self.send.on_invoke = self.keep_open
        def stop():
            if self.clock.now >= 0.4:
                raise StopRequested()
        with patch.object(v, 'check_stop', side_effect=stop):
            with self.assertRaises(StopRequested):
                self.submit()
        self.assertEqual(self.send.invoke_attempts, 1)

    def test_native_form_clicks_once_and_never_invokes(self):
        self.window.handle = 1234
        self.send.click_input = Mock(side_effect=lambda: setattr(self.window, 'visible', False))
        with patch.object(adding, 'activate_window') as focus, patch.object(v, '_native_click_ready') as ready, patch.object(v, '_native_window_exists', return_value=True):
            self.assertTrue(self.submit())
        focus.assert_called_once()
        ready.assert_called_once_with(self.window, self.send, 123)
        self.send.click_input.assert_called_once()
        self.assertEqual(self.send.invoke_attempts, 0)

    def test_native_failed_foreground_check_prevents_click(self):
        self.window.handle = 1234
        self.send.click_input = Mock()
        with patch.object(adding, 'activate_window'), patch.object(v, '_native_click_ready', side_effect=RuntimeError('blocked')):
            with self.assertRaisesRegex(RuntimeError, 'blocked'):
                self.submit()
        self.send.click_input.assert_not_called()
        self.assertEqual(self.send.invoke_attempts, 0)

    def test_unconfirmed_native_click_is_not_repeated(self):
        self.window.handle = 1234
        self.send.click_input = Mock()
        with patch.object(adding, 'activate_window'), patch.object(v, '_native_click_ready'), patch.object(v, '_native_window_exists', return_value=True):
            with self.assertRaises(v.SendUnconfirmed):
                self.submit(timeout=0)
        self.send.click_input.assert_called_once()
        self.assertEqual(self.send.invoke_attempts, 0)

    def test_direct_form_uses_same_native_click_path(self):
        window, _, _, send, _ = unconfigured_application()
        prepared = q.prepare_verification_form(Desktop(window), 123, MESSAGE, 0, expected_number=NUMBER)
        window.handle = 1234
        send.click_input = Mock(side_effect=lambda: setattr(window, 'visible', False))
        with patch.object(adding, 'activate_window'), patch.object(v, '_native_click_ready'), patch.object(v, '_native_window_exists', return_value=True):
            self.assertTrue(q.click_verified_send(prepared, 123, MESSAGE, expected_number=NUMBER))
        send.click_input.assert_called_once()
        self.assertEqual(send.invoke_attempts, 0)

    def test_wrong_target_is_never_closed(self):
        self.window.close = Mock()
        with self.assertRaisesRegex(RuntimeError, '号码不匹配'):
            v.close_submitted_form(self.window, 123, '123456789')
        self.window.close.assert_not_called()

    def test_matching_pending_form_can_be_closed_without_send(self):
        self.window.close = Mock(side_effect=lambda: setattr(self.window, 'visible', False))
        v.close_submitted_form(self.window, 123, NUMBER)
        self.window.close.assert_called_once()
        self.assertEqual(self.send.invoke_attempts, 0)

    def test_stale_uia_with_live_hwnd_is_not_treated_as_closed(self):
        self.window.handle = 1234
        stale = RuntimeError('stale wrapper')
        self.window.is_visible = Mock(side_effect=stale)
        with patch.object(v, '_native_window_exists', return_value=True), patch.object(v, '_is_unavailable_element', return_value=True), patch.object(v, '_refresh_form', return_value=self.window):
            with self.assertRaises(v.SendUnconfirmed):
                v._wait_form_closed(self.window, 1234, 0, process_id=123, expected_number=NUMBER)

    def test_transient_read_retries_without_retrying_bad_target(self):
        read = Mock(side_effect=[KeyError(None), 'ready'])
        self.assertEqual(v._retry_form_read(read), 'ready')
        self.assertEqual(read.call_count, 2)
        mismatch = Mock(side_effect=RuntimeError('目标号码不匹配'))
        with self.assertRaisesRegex(RuntimeError, '目标号码不匹配'):
            v._retry_form_read(mismatch)
        mismatch.assert_called_once()

    def test_native_pointer_guard_requires_same_pid_foreground_and_hit(self):
        import win32gui
        import win32process
        self.window.handle = 1234
        self.window.rectangle = lambda: SimpleNamespace(left=0, top=0, right=500, bottom=400)
        self.send.rectangle = lambda: SimpleNamespace(left=100, top=200, right=200, bottom=240)
        with patch.object(win32gui, 'GetForegroundWindow', return_value=1234), patch.object(win32process, 'GetWindowThreadProcessId', return_value=(1,123)), patch.object(win32gui, 'WindowFromPoint', return_value=1234):
            v._native_click_ready(self.window, self.send, 123)
            with patch.object(win32gui, 'GetForegroundWindow', return_value=999):
                with self.assertRaisesRegex(RuntimeError, '前台'):
                    v._native_click_ready(self.window, self.send, 123)
            with patch.object(win32process, 'GetWindowThreadProcessId', return_value=(1,456)):
                with self.assertRaisesRegex(RuntimeError, '前台'):
                    v._native_click_ready(self.window, self.send, 123)
            with patch.object(win32gui, 'WindowFromPoint', return_value=999), patch.object(win32gui, 'GetAncestor', return_value=999):
                with self.assertRaisesRegex(RuntimeError, '遮挡'):
                    v._native_click_ready(self.window, self.send, 123)
            self.send.enabled = False
            with self.assertRaisesRegex(RuntimeError, '不可用'):
                v._native_click_ready(self.window, self.send, 123)


class WorkflowResilienceTests(unittest.TestCase):
    def setup_sends(self, harness):
        harness.search_number.side_effect = lambda *args: ('results', 1)
        harness.click_result_add_button.side_effect = lambda *args, **kwargs: True

    def test_uncertain_send_is_recorded_unmarked_and_next_number_runs(self):
        with WorkflowHarness([fake_window(441)], target=441) as h:
            self.setup_sends(h)
            h.click_send_button.side_effect = [v.SendUnconfirmed('待核对'), True, True]
            h.run()
            self.assertEqual(h.click_send_button.call_count, 3)
            first = h.mark_rows_processed.call_args_list[0]
            self.assertFalse(first.kwargs['mark_processed'])
            self.assertIn('待核对', first.kwargs['reason'])
            self.assertEqual(h.journal.clear.call_count, 2)
            self.assertEqual(h.close_submitted_form.call_count, 3)

    def test_restriction_stops_before_next_number(self):
        with WorkflowHarness([fake_window(441)], target=441) as h:
            self.setup_sends(h)
            h.click_send_button.side_effect = v.SendRejected('操作过于频繁', halt_run=True)
            with self.assertRaisesRegex(RuntimeError, '频率限制'):
                h.run()
            self.assertEqual(h.search_number.call_count, 1)
            h.journal.clear.assert_not_called()
            h.close_submitted_form.assert_not_called()

    def test_cannot_close_unconfirmed_form_stops_without_resending(self):
        with WorkflowHarness([fake_window(441)], target=441) as h:
            self.setup_sends(h)
            h.click_send_button.side_effect = v.SendUnconfirmed('待核对')
            h.close_submitted_form.side_effect = RuntimeError('still open')
            with self.assertRaisesRegex(RuntimeError, '无法安全关闭'):
                h.run()
            self.assertEqual(h.click_send_button.call_count, 1)
            h.journal.clear.assert_not_called()


if __name__ == '__main__':
    unittest.main()
