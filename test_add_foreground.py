"""Foreground failures and recovery, with simulated windows only."""
import io
from contextlib import ExitStack, redirect_stdout
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import open_qq_add_friend as adding
from task_control import StopRequested
from test_add_account_detection import Clock, own_header, main_window, desktop


class ForegroundTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.stack.enter_context(redirect_stdout(io.StringIO()))
        self.clock = Clock()
        self.stack.enter_context(patch.object(adding, 'time', self.clock))
        self.window = Mock(handle=123)
        self.gui = Mock()
        self.gui.IsWindow.return_value = True
        self.gui.IsWindowVisible.return_value = True
        self.gui.IsIconic.return_value = False
        self.gui.GetForegroundWindow.return_value = 999
        self.con = SimpleNamespace(SW_RESTORE=9, SW_MINIMIZE=6)

    def focus_target(self):
        self.gui.GetForegroundWindow.return_value = self.window.handle

    def activate(self):
        adding.activate_window(self.window, self.con, self.gui)

    def test_zero_code_error_recovers_via_fallback(self):
        self.gui.SetForegroundWindow.side_effect = OSError(0, 'SetForegroundWindow', 'No error message is available')
        self.window.set_focus.side_effect = self.focus_target
        self.activate()
        self.window.set_focus.assert_called_once()
        self.assertEqual(self.clock.now, 100)

    def test_setforeground_can_raise_after_succeeding(self):
        def activate_then_raise(handle):
            self.focus_target()
            raise OSError(0, 'SetForegroundWindow')
        self.gui.SetForegroundWindow.side_effect = activate_then_raise
        self.activate()
        self.window.set_focus.assert_not_called()

    def test_silent_foreground_failure_also_uses_fallback(self):
        self.window.set_focus.side_effect = self.focus_target
        self.activate()
        self.window.set_focus.assert_called_once()

    def test_bring_to_top_failure_does_not_skip_setforeground(self):
        self.gui.BringWindowToTop.side_effect = OSError('denied')
        self.gui.SetForegroundWindow.side_effect = lambda handle: self.focus_target()
        self.activate()
        self.gui.SetForegroundWindow.assert_called_once_with(123)
        self.window.set_focus.assert_not_called()

    def test_second_round_recovers_after_both_activation_methods_fail(self):
        self.gui.SetForegroundWindow.side_effect = OSError(0, 'denied')
        def fallback():
            if self.window.set_focus.call_count == 1:
                raise RuntimeError('temporary focus error')
            self.focus_target()
        self.window.set_focus.side_effect = fallback
        self.activate()
        self.assertEqual(self.window.set_focus.call_count, 2)
        self.assertEqual(self.gui.SetForegroundWindow.call_count, 2)

    def test_delayed_foreground_is_polled_before_retry(self):
        self.gui.GetForegroundWindow.side_effect = lambda: 123 if self.clock.now >= 100.5 else 999
        self.activate()
        self.gui.SetForegroundWindow.assert_called_once()
        self.assertGreaterEqual(self.clock.now, 100.5)
        self.assertLess(self.clock.now, 101)

    def test_already_foreground_needs_no_activation(self):
        self.focus_target()
        self.activate()
        self.gui.SetForegroundWindow.assert_not_called()
        self.window.set_focus.assert_not_called()

    def test_hidden_or_minimized_window_is_restored(self):
        for iconic, visible in [(True, True), (False, False)]:
            with self.subTest(iconic=iconic, visible=visible):
                self.gui.reset_mock()
                self.gui.IsIconic.return_value = iconic
                self.gui.IsWindowVisible.return_value = visible
                self.gui.GetForegroundWindow.return_value = 999
                def restore(handle, mode):
                    self.gui.IsIconic.return_value = False
                    self.gui.IsWindowVisible.return_value = True
                self.gui.ShowWindow.side_effect = restore
                self.window.set_focus.side_effect = self.focus_target
                self.activate()
                self.gui.ShowWindow.assert_called_once_with(123, 9)

    def test_wrong_foreground_fails_after_bounded_retries(self):
        error = OSError(0, 'SetForegroundWindow')
        self.gui.SetForegroundWindow.side_effect = error
        with self.assertRaisesRegex(RuntimeError, '尝试 3 次.*前台') as caught:
            self.activate()
        self.assertEqual(self.gui.SetForegroundWindow.call_count, 3)
        self.assertEqual(self.window.set_focus.call_count, 3)
        self.assertIs(caught.exception.__cause__, error)
        self.assertLess(self.clock.now - 100, 4)

    def test_same_process_other_window_is_not_accepted(self):
        self.gui.GetForegroundWindow.return_value = 124
        with self.assertRaisesRegex(RuntimeError, '操作错账号'):
            self.activate()

    def test_invalid_handle_never_calls_focus(self):
        self.gui.IsWindow.return_value = False
        with self.assertRaisesRegex(RuntimeError, '窗口已失效'):
            self.activate()
        self.gui.SetForegroundWindow.assert_not_called()
        self.window.set_focus.assert_not_called()

    def test_stop_is_not_swallowed(self):
        self.gui.SetForegroundWindow.side_effect = StopRequested()
        with self.assertRaises(StopRequested):
            self.activate()
        self.window.set_focus.assert_not_called()

    def test_stop_during_polling_prevents_next_attempt(self):
        def check_stop():
            if self.clock.now >= 100.2:
                raise StopRequested()
        with patch.object(adding, 'check_stop', side_effect=check_stop):
            with self.assertRaises(StopRequested):
                self.activate()
        self.gui.SetForegroundWindow.assert_called_once()

    def test_mouse_failsafe_in_fallback_is_not_swallowed(self):
        failsafe = type('FailSafeException', (Exception,), {})
        self.window.set_focus.side_effect = failsafe('abort')
        with self.assertRaises(failsafe):
            self.activate()
        self.gui.SetForegroundWindow.assert_called_once()

    def test_avatar_is_never_clicked_when_foreground_cannot_be_confirmed(self):
        header, avatar = own_header()
        window = main_window(header)
        mouse = Mock()
        with patch.object(adding, 'find_own_profile_account', return_value=None), \
             patch.object(adding, 'own_profile_roots', return_value=[]):
            with self.assertRaisesRegex(RuntimeError, '尝试 3 次'):
                adding.read_qq_account_from_avatar(window, mouse, self.con, self.gui, desktop(window))
        avatar.click_input.assert_not_called()
        mouse.moveTo.assert_not_called()

    def test_avatar_read_continues_after_activation_fallback(self):
        header, avatar = own_header()
        window = main_window(header)
        self.window = window
        self.gui.SetForegroundWindow.side_effect = OSError(0, 'SetForegroundWindow')
        window.set_focus.side_effect = self.focus_target
        with patch.object(adding, 'find_own_profile_account', side_effect=lambda *args: '123456789' if avatar.click_input.called else None), \
             patch.object(adding, 'own_profile_roots', return_value=[]):
            account = adding.read_qq_account_from_avatar(window, Mock(), self.con, self.gui, desktop(window))
        self.assertEqual(account, '123456789')
        avatar.click_input.assert_called_once()
        window.set_focus.assert_called_once()


if __name__ == '__main__':
    unittest.main()
