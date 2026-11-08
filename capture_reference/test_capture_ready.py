"""Readiness regressions with fake windows/clocks and isolated settings only."""
from contextlib import ExitStack, redirect_stdout
import io
import json
import os
import tempfile
import unittest
from unittest.mock import Mock, patch

import gui_config
import zq_qq_001 as capture
from task_control import StopRequested
from test_capture_detection import Clock, Control
from test_capture_navigation import member_scene, FakeScroll, WORK_AREA


class ReadyWaitTests(unittest.TestCase):
    def test_immediate_success_never_sleeps(self):
        clock = Clock()
        with patch.object(capture, 'time', clock):
            self.assertEqual(capture.wait_for_result(lambda: 'ready', 2), 'ready')
        self.assertEqual(clock.sleeps, [])

    def test_delayed_success_returns_without_remainder_wait(self):
        clock = Clock()
        with patch.object(capture, 'time', clock):
            result = capture.wait_for_result(Mock(side_effect=[None, None, 'ready']), 2, 0.25)
        self.assertEqual(result, 'ready')
        self.assertEqual(clock.sleeps, [0.25, 0.25])

    def test_timeout_does_not_increase_original_budget(self):
        clock = Clock()
        with patch.object(capture, 'time', clock):
            self.assertIsNone(capture.wait_for_result(lambda: None, 1, 0.3))
        self.assertAlmostEqual(sum(clock.sleeps), 1)

    def test_stop_during_poll_is_not_swallowed(self):
        clock = Clock()
        with patch.object(capture, 'time', clock), patch.object(
                capture, 'check_stop', side_effect=[None, StopRequested()]):
            with self.assertRaises(StopRequested):
                capture.wait_for_result(lambda: None, 2)
        self.assertEqual(clock.sleeps, [capture.PROFILE_POLL_INTERVAL])

    def test_old_mode_keeps_saved_delay(self):
        clock = Clock()
        with patch.object(capture, 'time', clock), patch.object(capture, 'CONTINUE_WHEN_READY', False):
            self.assertTrue(capture.after_ui_action(lambda: True, 1))
        self.assertEqual(clock.sleeps, [1])

    def test_legacy_settings_keep_all_timing_values(self):
        with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ, {'QQ_ASSISTANT_STATE_DIR': folder}):
            old = gui_config.defaults('capture')
            old.pop('CONTINUE_WHEN_READY')
            old.update(WAIT_SECONDS=1.7, MEMBER_CLICK_DELAY=1.3, PROFILE_READ_TIMEOUT=2.8,
                       PROFILE_POLL_INTERVAL=0.35, MAX_NUMBER_RECOGNITION_ATTEMPTS=4)
            gui_config.config_path('capture').write_text(json.dumps(old), encoding='utf-8')
            config = gui_config.load_config('capture')
            self.assertTrue(config.pop('CONTINUE_WHEN_READY'))
            self.assertEqual(config, old)
            gui_config.save_config('capture', dict(config, CONTINUE_WHEN_READY=False))
            self.assertFalse(gui_config.load_config('capture')['CONTINUE_WHEN_READY'])

    def test_member_panel_waits_for_rows_not_only_container(self):
        main, panel, rows = member_scene(['测试成员'])
        clock = Clock()
        with patch.object(capture, 'time', clock), patch.object(
                capture, 'find_main_qq_window', return_value=main), patch.object(
                capture, 'collect_member_items', side_effect=[[], rows]):
            self.assertIs(capture.wait_for_member_panel(object(), WORK_AREA), main)
        self.assertEqual(clock.sleeps, [0.15])

    def test_member_panel_timeout_keeps_interface_wait_value(self):
        clock = Clock()
        with patch.object(capture, 'time', clock), patch.object(
                capture, 'find_main_qq_window', return_value=None), patch.object(capture, 'WAIT_SECONDS', 1):
            self.assertIsNone(capture.wait_for_member_panel(object(), WORK_AREA))
        self.assertAlmostEqual(sum(clock.sleeps), 1)

    def test_blank_virtual_panel_refreshes_without_moving_or_sleeping(self):
        main, panel, rows = member_scene(['甲'])
        panel.iface_scroll = FakeScroll(percent=30)
        clock = Clock()
        with patch.object(capture, 'time', clock), patch.object(
                capture, 'find_main_qq_window', return_value=main), patch.object(
                capture, 'collect_member_items', side_effect=[[], rows]):
            self.assertIs(capture.wait_for_member_panel(object(), WORK_AREA), main)
        self.assertEqual(panel.iface_scroll.calls, [(-1.0, 30.01), (-1.0, 30.0)])
        self.assertEqual(clock.sleeps, [])

    def test_scroll_success_does_not_wait_after_new_rows(self):
        main, panel, rows = member_scene()
        panel.iface_scroll = FakeScroll(callback=lambda: setattr(rows[-1], 'name', '新页成员'))
        clock = Clock()
        with patch.object(capture, 'time', clock), patch.object(capture, 'find_main_qq_window', return_value=main):
            self.assertTrue(capture.scroll_member_list(object(), Mock(), WORK_AREA)[0])
        self.assertEqual(clock.sleeps, [])

    def test_reset_scroll_position_does_not_accept_stale_group_rows(self):
        main, pane = object(), Mock()
        pane.descendants.return_value = []
        pane.iface_scroll = FakeScroll(percent=50)
        old = [('旧页', Control('旧页'))]
        new = [('首页', Control('首页'))]
        clock = Clock()
        with patch.object(capture, 'time', clock), patch.object(
                capture, 'find_main_qq_window', return_value=main), patch.object(
                capture, 'get_contact_group_list', return_value=pane), patch.object(
                capture, 'collect_joined_group_rows', side_effect=[old, old, new]):
            self.assertIs(capture.reset_joined_groups_to_top(object(), main, Mock(), WORK_AREA), main)
        self.assertEqual(clock.sleeps, [0.05])

    def test_visible_top_marker_needs_no_scroll_or_extra_wait(self):
        main = object()
        marker = Control('置顶群聊 1', 'Button', rect=(56, 265, 306, 301))
        pane = Control('联系人列表', 'Pane', [marker], rect=(56, 265, 306, 900))
        rows = [('目标群', Control('目标群', 'Button', rect=(56, 445, 306, 509)))]
        mouse = Mock()
        clock = Clock()
        with patch.object(capture, 'time', clock), patch.object(
                capture, 'get_contact_group_list', return_value=pane), patch.object(
                capture, 'collect_joined_group_rows', return_value=rows):
            self.assertIs(capture.reset_joined_groups_to_top(object(), main, mouse, WORK_AREA), main)
        self.assertEqual(clock.sleeps, [])
        mouse.scroll.assert_not_called()

    def test_unchanged_number_is_not_accepted_as_new_profile(self):
        clock = Clock()
        old = {'controls': {}, 'windows': {}}
        with patch.object(capture, 'time', clock), patch.object(
                capture, 'collect_qq_snapshot', return_value=old), patch.object(
                capture, 'PROFILE_READ_TIMEOUT', 2):
            result = capture.extract_member_number_from_popup(
                Mock(), old, 1, 2, (0, 0, 100, 100), Mock(), Mock())
        self.assertIsNone(result[0])
        self.assertAlmostEqual(sum(clock.sleeps), 2)

    def test_two_successful_members_have_no_fixed_wait(self):
        for immediate in (True, False):
            with self.subTest(immediate=immediate), ExitStack() as stack:
                main, panel, rows = member_scene(['甲', '乙'])
                clock = Clock()
                stack.enter_context(redirect_stdout(io.StringIO()))
                stack.enter_context(patch.object(capture, 'time', clock))
                stack.enter_context(patch.object(capture, 'CONTINUE_WHEN_READY', immediate))
                stack.enter_context(patch.object(capture, 'MAX_MEMBERS_PER_GROUP', 2))
                stack.enter_context(patch.object(capture, 'MEMBER_CLICK_DELAY', 1))
                stack.enter_context(patch.object(capture, 'find_main_qq_window', return_value=main))
                stack.enter_context(patch.object(capture, 'collect_qq_snapshot', return_value={}))
                stack.enter_context(patch.object(capture, 'extract_member_number_from_popup',
                    side_effect=[('11111111', '', {}), ('22222222', '', {})]))
                saved = stack.enter_context(patch.object(capture, 'append_number', return_value=True))
                result = capture.click_non_admin_members(
                    main, Mock(), Mock(), WORK_AREA, Mock(), Mock(), {}, Mock(), Mock(), set(), {})
                self.assertEqual(result[3], 2)
                self.assertEqual([call.args[3] for call in saved.call_args_list], ['11111111', '22222222'])
                self.assertEqual(clock.sleeps, [] if immediate else [1, 1])


if __name__ == '__main__':
    unittest.main()
