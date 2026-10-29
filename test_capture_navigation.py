"""Navigation regressions using fake UIA controls only; never operates real QQ."""

from contextlib import ExitStack, redirect_stdout
import io
import unittest
from unittest.mock import Mock, patch

import zq_qq as capture
from test_capture_detection import Clock, Control, Rectangle, session_list, session_row, window


WORK_AREA = (0, 0, 1200, 800)


def member(name, top, bottom=None, left=900, right=1150):
    row = Control(name, 'ListItem', rect=(left, top, right, bottom or top + 40))
    row.click_input = Mock()
    return row


def member_scene(names=None, rect=(900, 100, 1150, 500)):
    names = names if names is not None else [f'成员{i}' for i in range(10)]
    rows = [member(name, rect[1] + index * 40, left=rect[0], right=rect[2])
            for index, name in enumerate(names)]
    panel = Control('群成员列表', 'Pane', rows, rect=rect)
    main = window(panel)
    main.handle = 42
    main.process_id = lambda: 1234
    return main, panel, rows


class FakeScroll:
    def __init__(self, *, percent=0.0, view_size=10.0, scrollable=True, callback=None, failure=None):
        self.CurrentVerticalScrollPercent = percent
        self.CurrentVerticalViewSize = view_size
        self.CurrentVerticallyScrollable = scrollable
        self.calls = []
        self.callback = callback
        self.failure = failure

    def SetScrollPercent(self, horizontal, vertical):
        self.calls.append((horizontal, vertical))
        if self.failure:
            raise self.failure
        self.CurrentVerticalScrollPercent = vertical
        if self.callback:
            self.callback()


class NavigationTests(unittest.TestCase):
    def setUp(self):
        self.output = io.StringIO()
        self.stdout = redirect_stdout(self.output)
        self.stdout.__enter__()
        self.addCleanup(self.stdout.__exit__, None, None, None)

    def test_already_open_group_is_never_clicked(self):
        main, header = object(), object()
        with patch.object(capture, 'find_main_qq_window', return_value=main), \
                patch.object(capture, 'get_group_header_button', return_value=header), \
                patch.object(capture, 'get_session_list') as sessions, \
                patch.object(capture, 'click_control_in_work_area') as click, \
                patch.object(capture, 'wait_for_group_header') as wait:
            self.assertEqual(capture.ensure_group_chat_open(object(), '目标群', WORK_AREA), (main, header))
        sessions.assert_not_called()
        click.assert_not_called()
        wait.assert_not_called()

    def test_switching_group_clicks_target_once(self):
        row = session_row('目标群')
        main = window(session_list(row))
        opened, header = object(), object()
        desktop = object()
        clock = Clock()
        with patch.object(capture, 'find_main_qq_window', return_value=main), \
                patch.object(capture, 'get_group_header_button', return_value=None), \
                patch.object(capture, 'click_control_in_work_area', return_value=True) as click, \
                patch.object(capture, 'wait_for_group_header', return_value=(opened, header)) as wait, \
                patch.object(capture, 'time', clock):
            self.assertEqual(capture.ensure_group_chat_open(desktop, '目标群', WORK_AREA), (opened, header))
        click.assert_called_once_with(row, WORK_AREA)
        wait.assert_called_once_with(desktop, '目标群')

    def test_selected_group_finishing_load_is_rechecked_before_click(self):
        row = session_row('目标群')
        main = window(session_list(row))
        header = object()
        with patch.object(capture, 'find_main_qq_window', return_value=main), \
                patch.object(capture, 'get_group_header_button', side_effect=[None, header]), \
                patch.object(capture, 'click_control_in_work_area') as click:
            self.assertEqual(capture.ensure_group_chat_open(object(), '目标群', WORK_AREA), (main, header))
        click.assert_not_called()

    def test_missing_target_fails_without_clicking_another_chat(self):
        main = window(session_list(session_row('其他群')))
        with patch.object(capture, 'find_main_qq_window', return_value=main), \
                patch.object(capture, 'get_group_header_button', return_value=None), \
                patch.object(capture, 'click_control_in_work_area') as click:
            with self.assertRaisesRegex(RuntimeError, '无法重新定位'):
                capture.ensure_group_chat_open(object(), '目标群', WORK_AREA)
        click.assert_not_called()


class MemberGeometryTests(unittest.TestCase):
    def test_virtual_rows_are_clipped_and_sorted_by_visible_panel(self):
        first, second = member('第一', 120), member('第二', 200)
        above, below = member('上方', 20), member('下方', 520)
        outside = member('旁边', 300, left=1200, right=1450)
        duplicate_nested_row = member('重复包装', 120)
        visible = Control('群成员列表', 'Pane', rect=(900, 100, 1150, 500))
        virtual = Control('群成员列表', 'Window',
                          [second, below, outside, first, above, duplicate_nested_row],
                          rect=(900, 100, 1150, 8000))
        self.assertEqual(capture.collect_member_items(virtual, visible, WORK_AREA), [first, second])

    def test_rows_outside_work_area_are_not_returned(self):
        main, panel, rows = member_scene(['第一', '第二', '第三', '第四'])
        self.assertEqual(capture.collect_member_items(panel, panel, (0, 0, 1200, 180)), rows[:2])

    def test_group_wrappers_are_excluded_but_text_rows_are_kept(self):
        text = Control('成员昵称', 'Text', rect=(950, 120, 1100, 140))
        row = Control('', 'Group', [text], rect=(900, 100, 1150, 160))
        wrapper = Control('', 'Group', [row], rect=(900, 100, 1150, 500))
        panel = Control('群成员列表', 'Pane', [wrapper], rect=(900, 100, 1150, 500))
        self.assertEqual(capture.collect_member_items(panel, panel, WORK_AREA), [row])

    def test_click_is_clipped_to_panel_and_work_area(self):
        row = member('边界成员', 80, bottom=140, left=850, right=1180)
        visible = Rectangle(900, 100, 1150, 500)
        self.assertTrue(capture.click_control_in_work_area(row, (920, 0, 1100, 800), visible))
        row.click_input.assert_called_once_with(coords=(160, 40))

    def test_fully_outside_panel_does_not_click(self):
        row = member('越界成员', 40, bottom=80)
        self.assertFalse(capture.click_control_in_work_area(row, WORK_AREA, Rectangle(900, 100, 1150, 500)))
        row.click_input.assert_not_called()


class MemberScrollingTests(unittest.TestCase):
    def run_scroll(self, main, mouse=None, work_area=WORK_AREA, extra=None):
        mouse = mouse if mouse is not None else Mock()
        clock = Clock()
        with ExitStack() as stack:
            stack.enter_context(redirect_stdout(io.StringIO()))
            stack.enter_context(patch.object(capture, 'find_main_qq_window', return_value=main))
            stack.enter_context(patch.object(capture, 'time', clock))
            stack.enter_context(patch.object(capture, 'WAIT_SECONDS', 0.1))
            if extra:
                stack.enter_context(extra)
            result = capture.scroll_member_list(object(), mouse, work_area)
        return result, mouse, clock

    def test_scroll_api_has_priority_and_keeps_overlap(self):
        main, panel, rows = member_scene()
        panel.iface_scroll = FakeScroll(callback=lambda: setattr(rows[-1], 'name', '下一页新成员'))
        result, mouse, _ = self.run_scroll(main)
        self.assertTrue(result[0])
        self.assertIn('滚动接口', result[1])
        self.assertEqual(len(panel.iface_scroll.calls), 1)
        horizontal, target = panel.iface_scroll.calls[0]
        self.assertEqual(horizontal, -1.0)
        self.assertAlmostEqual(target, (400 * 0.7) / (400 * (100 / 10 - 1)) * 100)
        self.assertLess(target, 400 / (400 * (100 / 10 - 1)) * 100)
        mouse.moveTo.assert_not_called()
        mouse.scroll.assert_not_called()

    def test_small_page_still_scrolls_and_obeys_member_page_size(self):
        main, panel, rows = member_scene(['甲', '乙', '丙'], rect=(900, 100, 1150, 220))
        panel.iface_scroll = FakeScroll(callback=lambda: setattr(rows[-1], 'name', '丁'))
        result, mouse, _ = self.run_scroll(main, extra=patch.object(capture, 'MEMBER_PAGE_SIZE', 1))
        self.assertTrue(result[0])
        self.assertAlmostEqual(panel.iface_scroll.calls[0][1], 40 / (120 * 9) * 100)
        mouse.scroll.assert_not_called()

    def test_changed_later_row_is_detected_when_first_three_stay_fixed(self):
        main, panel, rows = member_scene()
        before = [row.name for row in rows[:3]]
        panel.iface_scroll = FakeScroll(callback=lambda: setattr(rows[8], 'name', '变化成员'))
        result, _, _ = self.run_scroll(main)
        self.assertTrue(result[0])
        self.assertEqual([row.name for row in rows[:3]], before)

    def test_api_failure_uses_single_notch_wheel_inside_clipped_panel(self):
        main, panel, rows = member_scene()
        panel.iface_scroll = FakeScroll(failure=RuntimeError('unsupported'))
        mouse = Mock()
        mouse.scroll.side_effect = lambda *args, **kwargs: setattr(rows[5], 'name', '新成员')
        result, mouse, _ = self.run_scroll(main, mouse, work_area=(950, 150, 1120, 450))
        self.assertTrue(result[0])
        self.assertIn('滚轮', result[1])
        args, kwargs = mouse.scroll.call_args
        self.assertEqual(args, (-1,))
        self.assertTrue(950 <= kwargs['x'] < 1120)
        self.assertTrue(150 <= kwargs['y'] < 450)
        self.assertEqual(mouse.scroll.call_count, 1)

    def test_missing_api_falls_back_to_wheel(self):
        main, panel, rows = member_scene()
        mouse = Mock()
        mouse.scroll.side_effect = lambda *args, **kwargs: setattr(rows[-1], 'name', '新成员')
        result, _, _ = self.run_scroll(main, mouse)
        self.assertTrue(result[0])
        self.assertEqual(mouse.scroll.call_count, 1)

    def test_no_change_is_bounded_to_three_attempts(self):
        main, panel, rows = member_scene()
        panel.iface_scroll = FakeScroll()
        result, mouse, clock = self.run_scroll(main)
        self.assertFalse(result[0])
        self.assertIn('连续 3 次', result[1])
        self.assertEqual(len(panel.iface_scroll.calls), 1)
        self.assertEqual(mouse.scroll.call_count, 2)
        self.assertLess(clock.now - 100, 4)

    def test_first_no_change_then_wheel_success(self):
        main, panel, rows = member_scene()
        panel.iface_scroll = FakeScroll()
        mouse = Mock()
        mouse.scroll.side_effect = lambda *args, **kwargs: setattr(rows[-1], 'name', '新成员')
        result, _, _ = self.run_scroll(main, mouse)
        self.assertTrue(result[0])
        self.assertEqual(len(panel.iface_scroll.calls), 1)
        self.assertEqual(mouse.scroll.call_count, 1)

    def test_bottom_and_fully_visible_list_never_scroll(self):
        for options in ({'percent': 100.0}, {'scrollable': False}):
            with self.subTest(options=options):
                main, panel, rows = member_scene()
                panel.iface_scroll = FakeScroll(**options)
                result, mouse, _ = self.run_scroll(main)
                self.assertFalse(result[0])
                self.assertEqual(panel.iface_scroll.calls, [])
                mouse.scroll.assert_not_called()

    def test_target_percent_is_capped_at_bottom(self):
        main, panel, rows = member_scene()
        panel.iface_scroll = FakeScroll(percent=99.0, callback=lambda: setattr(rows[-1], 'name', '末页成员'))
        result, _, _ = self.run_scroll(main)
        self.assertTrue(result[0])
        self.assertEqual(panel.iface_scroll.calls[0][1], 100.0)

    def test_temporary_empty_list_is_polled_until_rows_return(self):
        main, panel, before = member_scene()
        after = [member(row.name, row.rectangle().top) for row in before]
        after[-1].name = '新页成员'
        panel.iface_scroll = FakeScroll()
        with patch.object(capture, 'collect_member_items', side_effect=[before, [], after]) as collect:
            result, mouse, clock = self.run_scroll(main)
        self.assertTrue(result[0])
        self.assertEqual(collect.call_count, 3)
        self.assertGreaterEqual(len(clock.sleeps), 2)
        mouse.scroll.assert_not_called()

    def test_offscreen_panel_never_sends_mouse_input(self):
        main, panel, rows = member_scene()
        result, mouse, _ = self.run_scroll(main, work_area=(0, 0, 800, 800))
        self.assertFalse(result[0])
        mouse.moveTo.assert_not_called()
        mouse.scroll.assert_not_called()


class MemberProcessingTests(unittest.TestCase):
    def run_processing(self, pages, limit):
        main, panel, _ = member_scene(pages[0])
        mouse = Mock()
        page_index = 0
        clicked = []
        numbers = iter(str(100000 + index) for index in range(200))

        def click(row, work_area, clip_rect=None):
            clicked.append(row.name)
            self.assertEqual(clip_rect, panel.rectangle())
            return True

        def advance(*args):
            nonlocal page_index
            page_index += 1
            if page_index >= len(pages):
                return False, '已到底'
            panel._children = [member(name, 100 + index * 40) for index, name in enumerate(pages[page_index])]
            return True, '已推进'

        with ExitStack() as stack:
            stack.enter_context(redirect_stdout(io.StringIO()))
            stack.enter_context(patch.object(capture, 'find_main_qq_window', return_value=main))
            stack.enter_context(patch.object(capture, 'time', Clock()))
            stack.enter_context(patch.object(capture, 'MAX_MEMBER_PAGE_TURNS', limit))
            stack.enter_context(patch.object(capture, 'collect_qq_snapshot', return_value={'controls': {}, 'windows': {}}))
            stack.enter_context(patch.object(capture, 'click_control_in_work_area', side_effect=click))
            stack.enter_context(patch.object(capture, 'extract_member_number_from_popup',
                                            side_effect=lambda *args: (next(numbers), '模拟号码', {})))
            save = stack.enter_context(patch.object(capture, 'append_number', return_value=True))
            scroll = stack.enter_context(patch.object(capture, 'scroll_member_list', side_effect=advance))
            result = capture.click_non_admin_members(main, mouse, object(), WORK_AREA, object(), object(),
                                                     {}, object(), object(), set())
        return result, clicked, scroll.call_count, save.call_count

    def test_fewer_than_thirteen_members_continue_to_final_short_page(self):
        result, clicked, scrolls, saves = self.run_processing([['甲', '乙'], ['丙']], limit=2)
        self.assertEqual(clicked, ['甲', '乙', '丙'])
        self.assertEqual(result, (3, 0, 1, 3, 0, 0))
        self.assertEqual(scrolls, 2)
        self.assertEqual(saves, 3)

    def test_same_named_members_on_same_and_later_pages_are_all_processed(self):
        result, clicked, scrolls, saves = self.run_processing([['同名', '同名'], ['同名']], limit=1)
        self.assertEqual(clicked, ['同名', '同名', '同名'])
        self.assertEqual(result[0], 3)
        self.assertEqual(saves, 3)

    def test_zero_page_limit_processes_first_page_without_scroll(self):
        result, clicked, scrolls, saves = self.run_processing([['甲'], ['乙']], limit=0)
        self.assertEqual(clicked, ['甲'])
        self.assertEqual(result[2], 0)
        self.assertEqual(scrolls, 0)

    def test_page_limit_counts_only_successful_advances(self):
        result, clicked, scrolls, saves = self.run_processing([['甲'], ['乙'], ['丙'], ['丁']], limit=2)
        self.assertEqual(clicked, ['甲', '乙', '丙'])
        self.assertEqual(result[2], 2)
        self.assertEqual(scrolls, 2)

    def test_admins_and_named_exclusions_remain_skipped(self):
        result, clicked, scrolls, saves = self.run_processing([['群主 张三', '管理员 李四', 'Q群管家', '普通成员']], limit=0)
        self.assertEqual(clicked, ['普通成员'])
        self.assertEqual(result, (1, 3, 0, 1, 0, 0))


if __name__ == '__main__':
    unittest.main()
