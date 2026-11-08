"""Deterministic profile delays and retries; no real QQ or mouse actions."""
import io
import unittest
from contextlib import ExitStack, redirect_stdout
from unittest.mock import Mock, patch

import open_qq_add_friend as adding
from task_control import StopRequested
from test_add_account_detection import Clock, own_header, main_window, desktop


class AvatarRetryTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.stack.enter_context(redirect_stdout(io.StringIO()))
        self.clock = Clock()
        header, self.avatar = own_header()
        self.window = main_window(header)
        self.screen = desktop(self.window)
        self.mouse = Mock()
        self.stack.enter_context(patch.object(adding, 'time', self.clock))
        self.stack.enter_context(patch.object(adding, 'activate_window'))
        self.stack.enter_context(patch.object(adding, 'AVATAR_TIMEOUT_SECONDS', 5))
        self.stack.enter_context(patch.object(adding, 'PAGE_TIMEOUT_SECONDS', 20))
        self.roots = self.stack.enter_context(patch.object(adding, 'own_profile_roots', return_value=[]))
        self.identity = self.stack.enter_context(patch.object(adding, 'find_own_profile_account', return_value=None))

    def read(self):
        return adding.read_qq_account_from_avatar(self.window, self.mouse, Mock(), Mock(), self.screen)

    def test_second_click_recovers_after_first_click_and_hover_did_nothing(self):
        self.identity.side_effect = lambda *args: '123456789' if self.avatar.click_input.call_count >= 2 else None
        self.assertEqual(self.read(), '123456789')
        self.assertEqual(self.avatar.click_input.call_count, 2)
        self.assertEqual(self.mouse.moveTo.call_count, 2)
        self.assertGreaterEqual(self.clock.now - 100, 25)

    def test_long_hover_succeeds_after_eighteen_seconds_without_another_click(self):
        hovered = []
        def move(*args, **kwargs):
            if self.mouse.moveTo.call_count % 2 == 0:
                hovered.append(self.clock.now)
        self.mouse.moveTo.side_effect = move
        self.identity.side_effect = lambda *args: ('123456789' if hovered and self.clock.now - hovered[-1] >= 18 else None)
        self.assertEqual(self.read(), '123456789')
        self.assertEqual(self.avatar.click_input.call_count, 1)
        self.assertEqual(self.mouse.moveTo.call_count, 2)
        self.assertGreaterEqual(self.clock.now - hovered[-1], 18)

    def test_later_hover_gets_longer_wait_and_succeeds_after_twenty_two_seconds(self):
        hovered = []
        self.mouse.moveTo.side_effect = lambda *args, **kwargs: hovered.append(self.clock.now) if self.mouse.moveTo.call_count % 2 == 0 else None
        self.identity.side_effect = lambda *args: ('123456789' if len(hovered) >= 2 and self.clock.now - hovered[-1] >= 22 else None)
        self.assertEqual(self.read(), '123456789')
        self.assertEqual(self.avatar.click_input.call_count, 2)
        self.assertEqual(len(hovered), 2)

    def test_loading_card_gets_twenty_seconds_and_is_not_toggled_closed(self):
        self.roots.side_effect = lambda *args: [object()] if self.clock.now >= 104 else []
        self.identity.side_effect = lambda *args: '123456789' if self.clock.now >= 118 else None
        self.assertEqual(self.read(), '123456789')
        self.avatar.click_input.assert_called_once()
        self.mouse.moveTo.assert_not_called()

    def test_existing_loading_card_can_finish_after_twenty_five_seconds(self):
        self.roots.return_value = [object()]
        self.identity.side_effect = lambda *args: '123456789' if self.clock.now >= 125 else None
        self.assertEqual(self.read(), '123456789')
        self.avatar.click_input.assert_not_called()
        self.mouse.moveTo.assert_not_called()

    def test_missing_avatar_retries_until_account_bar_is_ready(self):
        with patch.object(adding, 'find_own_avatar', side_effect=[None, None, self.avatar]):
            self.identity.side_effect = lambda *args: '123456789' if self.avatar.click_input.called else None
            self.assertEqual(self.read(), '123456789')
        self.assertEqual(self.clock.now, 102)

    def test_click_exception_is_followed_by_hover_recovery(self):
        self.avatar.click_input.side_effect = RuntimeError('stale UIA element')
        self.identity.side_effect = lambda *args: '123456789' if self.mouse.moveTo.call_count == 2 else None
        self.assertEqual(self.read(), '123456789')
        self.assertEqual(self.mouse.moveTo.call_count, 2)

    def test_all_attempts_fail_with_bounded_clicks_and_hover_waits(self):
        with self.assertRaisesRegex(RuntimeError, '4 轮点击、悬停'):
            self.read()
        self.assertEqual(self.avatar.click_input.call_count, 4)
        self.assertEqual(self.mouse.moveTo.call_count, 8)
        self.assertGreaterEqual(self.clock.now - 100, 4 * 5 + 20 + 23 + 26 + 29)
        self.assertLess(self.clock.now - 100, 127)

    def test_stop_interrupts_wait_before_a_second_click(self):
        def stop():
            if self.clock.now >= 101:
                raise StopRequested()
        with patch.object(adding, 'check_stop', side_effect=stop):
            with self.assertRaises(StopRequested):
                self.read()
        self.avatar.click_input.assert_called_once()
        self.mouse.moveTo.assert_not_called()

    def test_mouse_failsafe_is_not_swallowed_or_retried(self):
        failsafe = type('FailSafeException', (Exception,), {})
        self.mouse.moveTo.side_effect = failsafe('abort')
        with self.assertRaises(failsafe):
            self.read()
        self.avatar.click_input.assert_called_once()

    def test_ambiguous_identity_is_never_retried_or_guessed(self):
        self.identity.side_effect = RuntimeError('多个不同的本人账号')
        with self.assertRaisesRegex(RuntimeError, '多个不同'):
            self.read()
        self.avatar.click_input.assert_not_called()
        self.mouse.moveTo.assert_not_called()

    def test_hover_reenters_avatar_without_clicking_another_control(self):
        adding.hover_own_avatar(self.window, self.avatar, self.mouse)
        moves = [call.args[:2] for call in self.mouse.moveTo.call_args_list]
        self.assertEqual(moves, [(633, 287), (536, 230)])
        self.mouse.click.assert_not_called()
        self.avatar.click_input.assert_not_called()


if __name__ == '__main__':
    unittest.main()
