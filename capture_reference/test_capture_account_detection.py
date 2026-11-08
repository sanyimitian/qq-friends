"""Account startup regressions using simulated controls; never operates QQ."""
import io
import unittest
from contextlib import redirect_stdout
from unittest.mock import Mock, patch

import zq_qq_001 as capture
from task_control import StopRequested
from test_add_account_detection import (Clock, Control, desktop, main_window,
                                        own_header, profile_card)


class CaptureProfileTests(unittest.TestCase):
    def test_nested_profile_wrappers_are_not_rescanned(self):
        inner = profile_card()
        outer = Control('资料卡', 'Window', [inner])
        outer.descendants = Mock(wraps=outer.descendants)
        main = main_window()
        self.assertEqual(capture.find_own_profile_account(main, desktop(main, outer)), '123456789')
        self.assertLessEqual(outer.descendants.call_count, 2)

    def test_web_document_profile_is_supported(self):
        document = Control('QQ 123456789 QQ空间 编辑资料 发消息', 'Document')
        main = main_window(document)
        self.assertEqual(capture.find_own_profile_account(main, desktop(main)), '123456789')

    def test_other_person_or_process_cannot_supply_account(self):
        main = main_window()
        for card in (profile_card(own=False), profile_card(process=99999),
                     profile_card(visible=False)):
            with self.subTest(card=card):
                self.assertIsNone(capture.find_own_profile_account(main, desktop(main, card)))

    def test_ambiguous_identity_stops(self):
        main = main_window()
        with self.assertRaises(RuntimeError):
            capture.find_own_profile_account(main, desktop(
                main, profile_card(), profile_card('987654321')))

    def test_loading_card_can_be_separate_or_embedded(self):
        main = main_window()
        self.assertTrue(capture.has_open_account_profile(
            main, desktop(main, profile_card(own=False))))
        embedded = main_window(Control('资料卡', 'Pane'))
        self.assertTrue(capture.has_open_account_profile(embedded, desktop(embedded)))

    def test_hidden_foreign_or_text_only_card_is_not_open(self):
        main = main_window(Control('资料卡', 'Text'))
        self.assertFalse(capture.has_open_account_profile(main, desktop(
            main, profile_card(visible=False), profile_card(process=99999))))

    def test_chat_document_is_not_treated_as_loading_profile(self):
        main = main_window(Control('聊天消息 QQ 987654321', 'Document'))
        self.assertFalse(capture.has_open_account_profile(main, desktop(main)))
        self.assertIsNone(capture.find_own_profile_account(main, desktop(main)))

    def test_main_window_uses_filtered_queries_only(self):
        main = main_window()
        original = main.descendants
        def filtered_only(**kwargs):
            self.assertTrue(kwargs, 'Do not scan the whole chat tree for account identity')
            return original(**kwargs)
        main.descendants = Mock(side_effect=filtered_only)
        self.assertEqual(capture.find_own_profile_account(
            main, desktop(main, profile_card())), '123456789')
        self.assertTrue(main.descendants.called)


class CaptureAccountStartupTests(unittest.TestCase):
    def setUp(self):
        header, self.avatar = own_header()
        self.main = main_window(header)
        self.screen = desktop(self.main)
        self.clock = Clock()
        self.started = self.clock.now
        for patcher in (patch.object(capture, 'time', self.clock),
                        patch.object(capture, 'AVATAR_TIMEOUT_SECONDS', 1.0),
                        patch.object(capture, 'activate_window')):
            patcher.start()
            self.addCleanup(patcher.stop)
        output = redirect_stdout(io.StringIO())
        output.__enter__()
        self.addCleanup(output.__exit__, None, None, None)

    def read(self):
        return capture.read_qq_account_from_avatar(
            self.main, Mock(), Mock(), Mock(), self.screen)

    def test_existing_readable_card_is_reused_without_click(self):
        self.screen = desktop(self.main, profile_card())
        self.assertEqual(self.read(), '123456789')
        self.avatar.click_input.assert_not_called()
        capture.activate_window.assert_not_called()

    def test_loading_existing_card_is_not_toggled_closed(self):
        self.screen = desktop(self.main, Control('资料卡', 'Window'))
        with patch.object(capture, 'find_own_profile_account',
                          side_effect=[None, None, '123456789']):
            self.assertEqual(self.read(), '123456789')
        self.avatar.click_input.assert_not_called()
        capture.activate_window.assert_not_called()

    def test_swallowed_first_click_is_retried_once(self):
        def on_click():
            if self.avatar.click_input.call_count == 2:
                self.screen.windows.return_value.append(profile_card())
        self.avatar.click_input.side_effect = on_click
        self.assertEqual(self.read(), '123456789')
        self.assertEqual(self.avatar.click_input.call_count, 2)
        self.assertEqual(capture.activate_window.call_count, 2)
        self.assertLessEqual(self.clock.now - self.started, 1.00001)

    def test_open_but_loading_card_is_never_double_clicked(self):
        loading = Control('资料卡', 'Window')
        def on_click():
            self.screen.windows.return_value.append(loading)
        self.avatar.click_input.side_effect = on_click
        def probe(*args):
            return '123456789' if self.clock.now - self.started >= 0.7 else None
        with patch.object(capture, 'find_own_profile_account', side_effect=probe):
            self.assertEqual(self.read(), '123456789')
        self.avatar.click_input.assert_called_once()

    def test_retry_does_not_extend_timeout_or_guess_account(self):
        with self.assertRaisesRegex(RuntimeError, '没有打开资料卡'):
            self.read()
        self.assertEqual(self.avatar.click_input.call_count, 2)
        self.assertAlmostEqual(self.clock.now - self.started, 1.0)

    def test_open_unreadable_profile_has_distinct_error(self):
        def on_click():
            self.screen.windows.return_value.append(profile_card(own=False))
        self.avatar.click_input.side_effect = on_click
        with self.assertRaisesRegex(RuntimeError, '未能从资料卡中确认'):
            self.read()
        self.avatar.click_input.assert_called_once()

    def test_existing_unreadable_card_times_out_without_clicking(self):
        self.screen = desktop(self.main, profile_card(own=False))
        with self.assertRaisesRegex(RuntimeError, '本人 QQ'):
            self.read()
        self.avatar.click_input.assert_not_called()
        self.assertAlmostEqual(self.clock.now - self.started, 1.0)

    def test_stop_interrupts_existing_card_wait(self):
        self.screen = desktop(self.main, Control('资料卡', 'Window'))
        with patch.object(capture, 'check_stop', side_effect=StopRequested):
            with self.assertRaises(StopRequested):
                self.read()
        self.avatar.click_input.assert_not_called()


class CaptureActivationTests(unittest.TestCase):
    def test_visible_window_must_also_be_foreground(self):
        main = main_window()
        gui = Mock()
        gui.IsIconic.return_value = False
        gui.IsWindowVisible.return_value = True
        gui.GetForegroundWindow.return_value = 999
        with patch.object(capture, 'time', Clock()):
            with self.assertRaises(RuntimeError):
                capture.activate_window(main, Mock(), gui)
        main.set_focus.assert_called_once()

    def test_focus_fallback_confirms_selected_window(self):
        main = main_window()
        gui = Mock()
        gui.IsIconic.return_value = False
        gui.IsWindowVisible.return_value = True
        gui.GetForegroundWindow.return_value = 999
        main.set_focus.side_effect = lambda: setattr(gui.GetForegroundWindow, 'return_value', main.handle)
        with patch.object(capture, 'time', Clock()):
            capture.activate_window(main, Mock(), gui)
        main.set_focus.assert_called_once()


if __name__ == '__main__':
    unittest.main()
