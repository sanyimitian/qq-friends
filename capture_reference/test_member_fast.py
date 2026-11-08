"""Profile-only reads: no full-chat traversal, no wait after a verified number."""
import unittest
from unittest.mock import Mock, patch
import zq_qq_001 as c
from test_capture_detection import Control, Clock


def snapshot(number, offset=0):
    controls = {}
    for text, bounds in [('QQ ', (100, 50, 124, 65)), (number, (123, 50, 200, 65))]:
        bounds = tuple(v + offset for v in bounds)
        controls[(1, text, bounds)] = dict(text=text, bounds=bounds, window_handle=1, control_type='Text')
    return dict(controls=controls, windows={1: dict(profile=True, bounds=(0, 0, 322, 420))})


class MemberFastTests(unittest.TestCase):
    def test_avatar_click_targets_left_image_inside_visible_row(self):
        row = Control('', 'Group', rect=(900, 100, 1100, 132))
        row.click_input = Mock()
        self.assertTrue(c.click_control_in_work_area(row, (0, 0, 1200, 800), avatar=True))
        row.click_input.assert_called_once_with(coords=(20, 16))

    def test_success_returns_before_first_sleep_or_diagnostic(self):
        clock = Clock()
        old = snapshot('11111111')
        with patch.object(c, 'time', clock), patch.object(c, 'collect_qq_snapshot', return_value=snapshot('22222222')), patch.object(c, 'snapshot_diagnostic') as diagnostic:
            result = c.extract_member_number_from_popup(Mock(), old, 1, 2, (0,0,1,1), Mock(), Mock())
        self.assertEqual(result[0], '22222222')
        self.assertEqual(clock.sleeps, [])
        diagnostic.assert_not_called()

    def test_repositioned_old_card_is_rejected(self):
        self.assertIsNone(c.select_profile_qq_number(snapshot('11111111', 50), snapshot('11111111')))

    def test_same_nickname_different_qq_is_accepted(self):
        self.assertEqual(c.select_profile_qq_number(snapshot('22222222'), snapshot('11111111')), '22222222')

    def test_combined_old_qq_label_is_not_accepted_after_layout_change(self):
        old = {'controls': {1: dict(text='QQ 11111111', bounds=(1,1,100,20),
                                   window_handle=1, control_type='Text')}, 'windows': {}}
        self.assertIsNone(c.select_profile_qq_number(snapshot('11111111', 50), old))

    def test_unlabeled_number_is_rejected(self):
        current = snapshot('22222222')
        current['controls'] = {k:v for k,v in current['controls'].items() if v['text'] != 'QQ '}
        self.assertIsNone(c.select_profile_qq_number(current, snapshot('11111111')))

    def test_no_new_number_retains_original_timeout(self):
        clock = Clock()
        old = snapshot('11111111')
        with patch.object(c, 'time', clock), patch.object(c, 'collect_qq_snapshot', return_value=old), patch.object(c, 'PROFILE_READ_TIMEOUT', 2):
            result = c.extract_member_number_from_popup(Mock(), old, 1, 2, (0,0,1,1), Mock(), Mock())
        self.assertIsNone(result[0])
        self.assertEqual(sum(clock.sleeps), 2)

    def test_browser_and_chat_are_never_traversed(self):
        card = Control('资料卡', 'Window', [Control('QQ ', 'Text'), Control('22222222', 'Text')])
        card.handle = 3; card.process_id = lambda: 9
        main = Mock(); main.handle = 2; main.process_id.return_value = 9
        main.window_text.return_value = 'QQ'
        browser = Mock(); browser.handle = 4; browser.process_id.return_value = 9
        browser.window_text.return_value = 'QQ(窗口)'
        desktop = Mock()
        with patch.object(c, 'member_profile_windows', return_value=[card]):
            result = c.collect_qq_snapshot(desktop, 9, 2, Mock(), Mock())
        self.assertEqual(len(result['controls']), 2)
        main.descendants.assert_not_called()
        browser.descendants.assert_not_called()
        desktop.windows.assert_not_called()

    def test_header_number_returns_before_loading_qq_space_text(self):
        label = Control('QQ', 'Text', rect=(100, 50, 124, 65))
        number = Control('22222222', 'Text', rect=(123, 50, 200, 65))
        slow = Control('QQ空间', 'Text')
        slow.window_text = Mock(side_effect=AssertionError('Unneeded provider call'))
        card = Control('资料卡', 'Window', [label, number, slow])
        card.handle = 3
        with patch.object(c, 'member_profile_windows', return_value=[card]):
            result = c.collect_qq_snapshot(Mock(), 9, 2, Mock(), Mock())
        self.assertEqual(len(result['controls']), 2)
        self.assertEqual(c.select_profile_qq_number(result, snapshot('11111111')), '22222222')
        slow.window_text.assert_not_called()

    def test_native_lookup_filters_process_and_avoids_desktop_uia_scan(self):
        gui = Mock()
        gui.EnumWindows.side_effect = lambda callback, unused: [callback(h, unused) for h in (2, 3, 4, 5)]
        gui.IsWindowVisible.return_value = True
        gui.GetWindowText.side_effect = lambda h: {2:'QQ', 3:'资料卡', 4:'QQ(窗口)', 5:'资料卡'}[h]
        desktop = Mock()
        with patch('win32process.GetWindowThreadProcessId', side_effect=lambda h:(1, 10 if h==5 else 9)):
            result = c.member_profile_windows(desktop, 9, 2, gui)
        desktop.window.assert_called_once_with(handle=3)
        desktop.windows.assert_not_called()
        self.assertEqual(len(result), 1)

if __name__ == '__main__':
    unittest.main()

