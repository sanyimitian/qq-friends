"""Pure UIA simulations: never opens QQ or sends mouse/keyboard input."""

import re
import unittest
from dataclasses import dataclass
from types import SimpleNamespace
from unittest.mock import patch

import zq_qq as capture


@dataclass(frozen=True)
class Rectangle:
    left: int
    top: int
    right: int
    bottom: int

    def width(self):
        return self.right - self.left

    def height(self):
        return self.bottom - self.top


class Control:
    def __init__(self, name='', kind='Group', children=(), rect=(100, 100, 400, 160)):
        self.name = name
        self.element_info = SimpleNamespace(control_type=kind)
        self._children = list(children)
        self._rectangle = Rectangle(*rect)

    def window_text(self):
        return self.name

    def children(self):
        return list(self._children)

    def rectangle(self):
        return self._rectangle

    def descendants(self, control_type=None):
        result = []
        for child in self._children:
            if control_type is None or child.element_info.control_type == control_type:
                result.append(child)
            result.extend(child.descendants(control_type=control_type))
        return result


class Clock:
    def __init__(self):
        self.now = 100.0
        self.sleeps = []

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        self.sleeps.append(seconds)
        self.now += seconds


def text_control(name, left=150, top=110, right=300, bottom=130):
    return Control(name, 'Text', rect=(left, top, right, bottom))


def session_row(name, top=100, kind='Group', extra=(), left=100, right=400):
    title = text_control(name, left=left + 50, top=top + 10, right=right - 60, bottom=top + 30)
    return Control(kind=kind, children=[title, *extra], rect=(left, top, right, top + 60))


def session_list(*rows):
    return Control('会话列表', 'List', children=rows, rect=(100, 100, 400, 400))


def window(*controls):
    return Control('QQ', 'Window', children=controls, rect=(0, 0, 1200, 800))


class SessionDetectionTests(unittest.TestCase):
    def test_nested_unread_badges_are_not_conversations(self):
        badge = Control(kind='Group', children=[text_control('99+', 350, 115, 375, 135)],
                        rect=(345, 110, 380, 140))
        row = session_row('测试群', extra=[badge])
        found = capture.collect_session_rows(session_list(row))
        self.assertEqual(found, [('测试群', row)])

    def test_numeric_and_badge_like_real_names_remain_valid(self):
        for name in ('123456', '99+', '0'):
            with self.subTest(name=name):
                row = session_row(name)
                self.assertEqual(capture.collect_session_rows(session_list(row)), [(name, row)])

    def test_offscreen_zero_height_and_outside_rows_are_excluded(self):
        visible = session_row('可见', 160)
        above = session_row('上方', 20)
        below = session_row('下方', 410)
        outside = session_row('旁边', 230, left=410, right=710)
        zero = Control(children=[text_control('零高度')], rect=(100, 100, 400, 100))
        self.assertEqual(capture.collect_session_rows(session_list(above, below, outside, zero, visible)),
                         [('可见', visible)])

    def test_rows_are_ordered_top_to_bottom_across_supported_types(self):
        first = session_row('第一', 100, kind='ListItem')
        middle = session_row('第二', 180, kind='DataItem')
        last = session_row('第三', 260)
        found = capture.collect_session_rows(session_list(last, first, middle))
        self.assertEqual(found, [('第一', first), ('第二', middle), ('第三', last)])

    def test_title_uses_position_instead_of_child_enumeration_order(self):
        title = text_control('真实群名', 150, 110, 280, 130)
        timestamp = text_control('12:34', 340, 110, 390, 130)
        preview = text_control('消息预览 (123)', 150, 140, 390, 155)
        row = Control(children=[preview, timestamp, title])
        self.assertEqual(capture.collect_session_rows(session_list(row)), [('真实群名', row)])

    def test_nested_same_row_is_not_counted_twice(self):
        nested = session_row('同一聊天')
        outer = session_row('同一聊天', extra=[nested])
        self.assertEqual(capture.collect_session_rows(session_list(outer)), [('同一聊天', outer)])


class HeaderDetectionTests(unittest.TestCase):
    def setUp(self):
        self.pattern = patch.object(capture, 'GROUP_HEADER_PATTERN', re.compile(r'\(\s*\d+\s*\)\s*$'))
        self.pattern.start()
        self.addCleanup(self.pattern.stop)

    def test_legacy_ascii_header(self):
        header = Control('测试群 (123)', 'Button')
        self.assertIs(capture.get_group_header_button(window(header), '测试群'), header)

    def test_fullwidth_count_and_parentheses_in_group_name(self):
        for name, title in [('测试群', '测试群（123）'), ('项目（讨论）群', '项目（讨论）群（123）')]:
            with self.subTest(title=title):
                header = Control(title, 'Button')
                self.assertIs(capture.get_group_header_button(window(header), name), header)

    def test_empty_button_with_separate_title_and_count_texts(self):
        header = Control('', 'Button', [text_control('测试群'), text_control('（123）')])
        self.assertIs(capture.get_group_header_button(window(header), '测试群'), header)

    def test_observed_qq_header_without_count_uses_group_toolbar(self):
        header = Control('LeLeAPI售后', 'Button', rect=(436, 77, 612, 105))
        group_app = Control('群应用', 'Button', rect=(1988, 69, 2043, 114))
        invite = Control('邀请加群', 'Button', rect=(2048, 69, 2092, 114))
        # Both actions exist in the observed QQ window; either is sufficient.
        for actions in ([group_app, invite], [group_app], [invite]):
            with self.subTest(actions=[action.name for action in actions]):
                main = Control('QQ', 'Window', [header, *actions], rect=(0, 0, 2200, 1400))
                self.assertIs(capture.get_group_header_button(main, 'LeLeAPI售后'), header)

    def test_countless_header_without_group_toolbar_is_not_a_group(self):
        header = Control('LeLeAPI售后', 'Button', rect=(436, 77, 612, 105))
        generic_action = Control('设置', 'Button', rect=(1988, 69, 2043, 114))
        self.assertIsNone(capture.get_group_header_button(window(header, generic_action), 'LeLeAPI售后'))

    def test_group_actions_in_message_area_do_not_validate_header(self):
        header = Control('LeLeAPI售后', 'Button', rect=(436, 77, 612, 105))
        for top, bottom in [(300, 345), (105, 150), (20, 77)]:
            with self.subTest(top=top, bottom=bottom):
                action = Control('群应用', 'Button', rect=(1988, top, 2043, bottom))
                self.assertIsNone(capture.get_group_header_button(window(header, action), 'LeLeAPI售后'))

    def test_group_actions_left_of_header_do_not_validate_header(self):
        header = Control('LeLeAPI售后', 'Button', rect=(436, 77, 612, 105))
        for left, right in [(100, 160), (580, 650)]:
            with self.subTest(left=left, right=right):
                action = Control('邀请加群', 'Button', rect=(left, 69, right, 114))
                self.assertIsNone(capture.get_group_header_button(window(header, action), 'LeLeAPI售后'))

    def test_group_toolbar_does_not_validate_different_chat_name(self):
        header = Control('LeLeAPI售后', 'Button', rect=(436, 77, 612, 105))
        action = Control('群应用', 'Button', rect=(1988, 69, 2043, 114))
        for name in ('其他群', 'LeLeAPI', 'LeLeAPI售后交流群'):
            with self.subTest(name=name):
                self.assertIsNone(capture.get_group_header_button(window(header, action), name))

    def test_nonbutton_or_zero_size_group_actions_are_not_evidence(self):
        header = Control('LeLeAPI售后', 'Button', rect=(436, 77, 612, 105))
        actions = [Control('群应用', 'Text', rect=(1988, 69, 2043, 114)),
                   Control('邀请加群', 'Button', rect=(2048, 69, 2048, 114)),
                   Control('群应用', 'Button', rect=(1988, 69, 2043, 69))]
        for action in actions:
            with self.subTest(kind=action.element_info.control_type, rect=action.rectangle()):
                self.assertIsNone(capture.get_group_header_button(window(header, action), 'LeLeAPI售后'))

    def test_private_chat_plain_header_and_message_text_are_not_groups(self):
        main = window(Control('张三', 'Button'), Control('张三 (123)', 'Text'),
                      Control('张三 (123)', 'Group'))
        self.assertIsNone(capture.get_group_header_button(main, '张三'))

    def test_different_group_and_invalid_counts_are_not_matches(self):
        for title in ('其他群 (123)', '测试群 (abc)', '测试群 (123) 新消息', '测试群'):
            with self.subTest(title=title):
                self.assertIsNone(capture.get_group_header_button(window(Control(title, 'Button')), '测试群'))

    def test_user_supplied_pattern_remains_supported(self):
        header = Control('测试群 [123]', 'Button')
        with patch.object(capture, 'GROUP_HEADER_PATTERN', re.compile(r'\[\d+\]$')):
            self.assertIs(capture.get_group_header_button(window(header), '测试群'), header)


class MemberPanelDetectionTests(unittest.TestCase):
    def test_new_legacy_and_custom_names_support_container_types(self):
        with patch.object(capture, 'MEMBER_LIST_TEXT', '自定义成员区域'):
            for name in ('成员列表', '群成员列表', '自定义成员区域'):
                for kind in ('Pane', 'Window', 'Group', 'List'):
                    with self.subTest(name=name, kind=kind):
                        panel = Control(name, kind, rect=(1915, 437, 2154, 1376))
                        main = window(panel)
                        self.assertIs(capture.get_member_list_window(main), panel)
                        self.assertTrue(capture.member_panel_is_open(main))

    def test_zero_size_panels_are_skipped_before_visible_panel(self):
        hidden_width = Control('群成员列表', 'Pane', rect=(1915, 437, 1915, 1376))
        hidden_height = Control('成员列表', 'Window', rect=(1915, 437, 2154, 437))
        visible = Control('群成员列表', 'Pane', rect=(1915, 437, 2154, 1376))
        self.assertFalse(capture.member_panel_is_open(window(hidden_width, hidden_height)))
        self.assertIs(capture.get_member_list_window(window(hidden_width, hidden_height, visible)), visible)

    def test_text_and_button_labels_do_not_count_as_member_panels(self):
        for kind in ('Text', 'Button', 'ListItem'):
            with self.subTest(kind=kind):
                fake = Control('群成员列表', kind, rect=(1915, 437, 2154, 1376))
                self.assertIsNone(capture.get_member_list_window(window(fake)))
                self.assertFalse(capture.member_panel_is_open(window(fake)))

    def test_unrelated_named_pane_does_not_count_as_member_panel(self):
        unrelated = Control('聊天记录', 'Pane', rect=(1915, 437, 2154, 1376))
        self.assertFalse(capture.member_panel_is_open(window(unrelated)))

    def test_observed_pane_and_list_are_used_as_scroll_container(self):
        for kind in ('Pane', 'List'):
            with self.subTest(kind=kind):
                unrelated = Control('侧边栏', 'Group', rect=(1700, 100, 2200, 1390))
                panel = Control('群成员列表', kind, rect=(1915, 437, 2154, 1376))
                main = Control('QQ', 'Window', [unrelated, panel], rect=(0, 0, 2200, 1400))
                self.assertIs(capture.get_visible_member_panel(main), panel)

    def test_legacy_virtual_window_uses_visible_group_geometry(self):
        visible = Control('成员列表', 'Group', rect=(1915, 437, 2154, 1376))
        virtual = Control('成员列表', 'Window', [visible], rect=(1915, 437, 2154, 8000))
        self.assertIs(capture.get_visible_member_panel(window(virtual)), visible)

    def test_nested_pane_is_preferred_over_virtual_window_geometry(self):
        visible = Control('群成员列表', 'Pane', rect=(1915, 437, 2154, 1376))
        virtual = Control('群成员列表', 'Window', [visible], rect=(1915, 437, 2154, 8000))
        self.assertIs(capture.get_visible_member_panel(window(virtual)), visible)


class HeaderPollingTests(unittest.TestCase):
    def test_delayed_window_and_header_are_retried(self):
        clock = Clock()
        desktop = object()
        loading = window(Control('加载中', 'Text'))
        header = Control('测试群 (123)', 'Button')
        loaded = window(header)
        with patch.object(capture, 'time', clock), patch.object(capture, 'WAIT_SECONDS', 0.1), \
                patch.object(capture, 'find_main_qq_window', side_effect=[None, loading, loaded]) as find:
            self.assertEqual(capture.wait_for_group_header(desktop, '测试群'), (loaded, header))
        self.assertEqual(find.call_count, 3)
        find.assert_called_with(desktop)
        self.assertEqual(clock.sleeps, [0.25, 0.25])

    def test_immediate_header_does_not_sleep(self):
        clock = Clock()
        header = Control('测试群 (123)', 'Button')
        loaded = window(header)
        with patch.object(capture, 'time', clock), \
                patch.object(capture, 'find_main_qq_window', return_value=loaded):
            self.assertEqual(capture.wait_for_group_header(object(), '测试群'), (loaded, header))
        self.assertEqual(clock.sleeps, [])

    def test_timeout_returns_latest_window_and_respects_wait_setting(self):
        for wait, timeout in [(0.1, 3.0), (2.0, 6.0)]:
            with self.subTest(wait=wait):
                clock = Clock()
                private_chat = window(Control('张三', 'Button'))
                with patch.object(capture, 'time', clock), patch.object(capture, 'WAIT_SECONDS', wait), \
                        patch.object(capture, 'find_main_qq_window', return_value=private_chat):
                    self.assertEqual(capture.wait_for_group_header(object(), '张三'), (private_chat, None))
                self.assertEqual(clock.now, 100.0 + timeout)
                self.assertEqual(sum(clock.sleeps), timeout)


if __name__ == '__main__':
    unittest.main()
