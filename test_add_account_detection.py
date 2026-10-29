"""Offline account-identity regressions; no QQ, mouse, or user data access."""

import io
import unittest
from contextlib import redirect_stdout
from dataclasses import dataclass
from types import SimpleNamespace
from unittest.mock import Mock, patch

import open_qq_add_friend as adding


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
    def __init__(self, name='', kind='Group', children=(),
                 rect=(0, 0, 100, 40), visible=True, enabled=True,
                 process=12824, handle=None):
        self.name = name
        self.element_info = SimpleNamespace(
            control_type=kind, name=name, process_id=process,
            automation_id='', class_name='',
        )
        self._children = list(children)
        self._parent = None
        for child in self._children:
            child._parent = self
        self._rectangle = Rectangle(*rect)
        self._visible = visible
        self._enabled = enabled
        self._process = process
        self.handle = handle if handle is not None else id(self)
        self.click_input = Mock()
        self.set_focus = Mock()

    def window_text(self):
        return self.name

    def children(self, control_type=None):
        return [child for child in self._children
                if control_type is None or child.element_info.control_type == control_type]

    def descendants(self, control_type=None):
        result = []
        for child in self._children:
            if control_type is None or child.element_info.control_type == control_type:
                result.append(child)
            result.extend(child.descendants(control_type=control_type))
        return result

    def parent(self):
        return self._parent

    def rectangle(self):
        return self._rectangle

    def is_visible(self):
        return self._visible

    def is_enabled(self):
        return self._enabled

    def process_id(self):
        return self._process


def avatar(name='本人昵称的头像', **kwargs):
    return Control(name, 'Button', rect=(520, 213, 553, 247), **kwargs)


def own_header(button=None, status_visible=True):
    button = button if button is not None else avatar()
    status = Control('在线状态', 'Button', rect=(564, 216, 594, 244),
                     visible=status_visible)
    header = Control(children=[button, status,
        Control('个性签名:未设置,点击编辑', 'Button', rect=(600, 216, 800, 244))],
        rect=(503, 203, 1536, 258))
    return header, button


def main_window(*controls, process=12824):
    return Control('QQ', 'Window', controls, rect=(428, 203, 1786, 1238),
                   process=process)


class Clock:
    def __init__(self):
        self.now = 100.0

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds


def profile_card(number='123456789', *, process=12824, combined=False,
                 own=True, visible=True, title='资料卡', extra=()):
    controls = [Control('本人昵称的头像', 'Button', rect=(570, 270, 620, 320))]
    if combined:
        controls.append(Control('QQ：' + number, 'Text', rect=(658, 335, 850, 357)))
    else:
        controls.extend([
            Control('QQ ', 'Text', rect=(658, 335, 690, 357)),
            Control(number, 'Text', rect=(689, 335, 800, 357)),
        ])
    if own:
        controls.append(Control('编辑资料', 'Button', rect=(658, 550, 760, 590)))
    controls.extend(extra)
    return Control(title, 'Window', controls, rect=(540, 260, 1050, 640),
                   process=process, visible=visible)


def desktop(*windows):
    return SimpleNamespace(windows=Mock(return_value=list(windows)))


class OwnAvatarTests(unittest.TestCase):
    def test_online_status_group_beats_earlier_chat_member_avatar(self):
        member = avatar('群员的头像')
        chat = Control('聊天内容', 'Pane', [member], rect=(800, 300, 1700, 1000))
        header, own = own_header()
        self.assertIs(adding.find_own_avatar(main_window(chat, header)), own)
        member.click_input.assert_not_called()
        own.click_input.assert_not_called()

    def test_configured_text_in_chat_does_not_beat_online_status_group(self):
        member = avatar('昵称')
        header, own = own_header()
        with patch.object(adding, 'AVATAR_BUTTON_TEXT', '昵称'), \
             patch.object(adding, 'AVATAR_AUTO_DETECT', True):
            self.assertIs(adding.find_own_avatar(main_window(
                Control('群成员列表', 'Pane', [member]), header)), own)

    def test_invisible_avatar_is_not_selected(self):
        hidden_header, hidden = own_header(avatar('隐藏的头像', visible=False))
        header, own = own_header()
        self.assertIs(adding.find_own_avatar(main_window(hidden_header, header)), own)
        hidden.click_input.assert_not_called()

    def test_disabled_avatar_is_not_selected(self):
        disabled_header, disabled = own_header(avatar('禁用的头像', enabled=False))
        header, own = own_header()
        self.assertIs(adding.find_own_avatar(main_window(disabled_header, header)), own)
        disabled.click_input.assert_not_called()

    def test_hidden_online_status_does_not_qualify_wrong_group(self):
        wrong_header, wrong = own_header(avatar('群员的头像'), status_visible=False)
        header, own = own_header()
        self.assertIs(adding.find_own_avatar(main_window(wrong_header, header)), own)
        wrong.click_input.assert_not_called()

    def test_ambiguous_account_headers_fail_before_any_click(self):
        first_header, first = own_header(avatar('甲的头像'))
        second_header, second = own_header(avatar('乙的头像'))
        with self.assertRaises(RuntimeError):
            adding.find_own_avatar(main_window(first_header, second_header))
        first.click_input.assert_not_called()
        second.click_input.assert_not_called()

    def test_legacy_header_uses_unique_top_left_avatar(self):
        own = avatar()
        member = Control('群员的头像', 'Button', rect=(1650, 450, 1680, 480))
        self.assertIs(adding.find_own_avatar(main_window(member, own)), own)

    def test_chat_avatar_alone_is_not_used_as_account_avatar(self):
        member = Control('群员的头像', 'Button', rect=(1650, 450, 1680, 480))
        self.assertIsNone(adding.find_own_avatar(main_window(member)))

    def test_multiple_legacy_header_candidates_are_not_guessed(self):
        self.assertIsNone(adding.find_own_avatar(main_window(avatar('甲的头像'), avatar('乙的头像'))))

    def test_explicit_name_works_with_automatic_detection_disabled(self):
        header, own = own_header(avatar('配置昵称'))
        with patch.object(adding, 'AVATAR_BUTTON_TEXT', '配置昵称'), \
             patch.object(adding, 'AVATAR_AUTO_DETECT', False):
            self.assertIs(adding.find_own_avatar(main_window(header)), own)


class OwnProfileTests(unittest.TestCase):
    def test_separate_same_process_card_supports_split_label(self):
        main = main_window()
        self.assertEqual(adding.find_own_profile_account(main, desktop(main, profile_card())),
                         '123456789')

    def test_separate_same_process_card_supports_combined_label(self):
        main = main_window()
        self.assertEqual(adding.find_own_profile_account(
            main, desktop(main, profile_card(combined=True))), '123456789')

    def test_different_process_card_does_not_supply_account(self):
        main = main_window()
        self.assertIsNone(adding.find_own_profile_account(
            main, desktop(main, profile_card(process=22334))))

    def test_hidden_card_does_not_supply_account(self):
        main = main_window()
        self.assertIsNone(adding.find_own_profile_account(
            main, desktop(main, profile_card(visible=False))))

    def test_other_persons_card_without_edit_button_is_not_own_account(self):
        main = main_window()
        self.assertIsNone(adding.find_own_profile_account(
            main, desktop(main, profile_card(own=False))))

    def test_hidden_edit_button_does_not_identify_own_card(self):
        main = main_window()
        card = profile_card(own=False, extra=[Control('编辑资料', 'Button', visible=False)])
        self.assertIsNone(adding.find_own_profile_account(main, desktop(main, card)))

    def test_chat_message_containing_qq_number_is_ignored(self):
        main = main_window(Control('我的QQ号是987654321', 'Text'))
        self.assertIsNone(adding.find_own_profile_account(main, desktop(main)))

    def test_profile_text_with_embedded_number_is_not_treated_as_account_label(self):
        main = main_window()
        card = Control('资料卡', 'Window', [Control('编辑资料', 'Button'),
            Control('联系我的QQ：987654321', 'Text')])
        self.assertIsNone(adding.find_own_profile_account(main, desktop(main, card)))

    def test_unrelated_numeric_field_below_qq_label_is_not_used(self):
        main = main_window()
        card = Control('资料卡', 'Window', [
            Control('编辑资料', 'Button', rect=(658, 550, 760, 590)),
            Control('QQ', 'Text', rect=(658, 335, 690, 357)),
            Control('987654321', 'Text', rect=(689, 435, 800, 457)),
        ])
        self.assertIsNone(adding.find_own_profile_account(main, desktop(main, card)))

    def test_numeric_sibling_left_of_qq_label_is_not_used(self):
        main = main_window()
        card = Control('资料卡', 'Window', [
            Control('编辑资料', 'Button', rect=(658, 550, 760, 590)),
            Control('QQ', 'Text', rect=(658, 335, 690, 357)),
            Control('987654321', 'Text', rect=(520, 335, 600, 357)),
        ])
        self.assertIsNone(adding.find_own_profile_account(main, desktop(main, card)))

    def test_hidden_number_is_not_used(self):
        main = main_window()
        card = profile_card()
        card.children()[2]._visible = False
        self.assertIsNone(adding.find_own_profile_account(main, desktop(main, card)))

    def test_nested_profile_wrapper_with_same_number_is_not_ambiguous(self):
        main = main_window()
        inner = profile_card()
        outer = Control('资料卡', 'Window', [inner], process=12824)
        self.assertEqual(adding.find_own_profile_account(main, desktop(main, outer)),
                         '123456789')

    def test_distinct_account_numbers_in_own_cards_raise(self):
        main = main_window()
        with self.assertRaises(RuntimeError):
            adding.find_own_profile_account(main, desktop(main,
                profile_card('123456789'), profile_card('987654321')))


class ReadAccountTests(unittest.TestCase):
    def test_existing_profile_reused_on_repeated_calls_without_clicking(self):
        header, own = own_header()
        main = main_window(header)
        screen = desktop(main, profile_card())
        mouse = Mock()
        with patch.object(adding, 'activate_window') as activate:
            for _ in range(2):
                self.assertEqual(adding.read_qq_account_from_avatar(
                    main, mouse, Mock(), Mock(), desktop=screen), '123456789')
        own.click_input.assert_not_called()
        activate.assert_not_called()
        mouse.press.assert_not_called()

    def test_delayed_profile_opens_avatar_once_and_polls(self):
        header, own = own_header()
        main = main_window(header)
        clock = Clock()
        mouse = Mock()
        screen = desktop(main)
        with patch.object(adding, 'activate_window') as activate, \
             patch.object(adding, 'find_own_profile_account',
                          side_effect=[None, None, '123456789']) as find_profile, \
             patch.object(adding, 'time', clock):
            self.assertEqual(adding.read_qq_account_from_avatar(
                main, mouse, Mock(), Mock(), desktop=screen), '123456789')
        self.assertGreaterEqual(find_profile.call_count, 3)
        own.click_input.assert_called_once()
        activate.assert_called_once()
        mouse.press.assert_not_called()

    def test_timeout_stops_instead_of_using_chat_number(self):
        header, own = own_header()
        main = main_window(header, Control('QQ：987654321', 'Text'))
        clock = Clock()
        with patch.object(adding, 'activate_window'), \
             patch.object(adding, 'AVATAR_TIMEOUT_SECONDS', 0.4), \
             patch.object(adding, 'time', clock):
            with self.assertRaises(RuntimeError):
                adding.read_qq_account_from_avatar(main, Mock(), Mock(), Mock(),
                                                  desktop=desktop(main))
        own.click_input.assert_called_once()


class MissingMainWindowTests(unittest.TestCase):
    def missing_window_error(self, process):
        dependencies = tuple(Mock() for _ in range(10))
        with patch.object(adding.sys, 'platform', 'win32'), \
             patch.object(adding, 'NUMBERS_PER_ACCOUNT', 3), \
             patch.object(adding, 'TARGET_PROCESS_ID', process), \
             patch.object(adding, 'load_dependencies', return_value=dependencies), \
             patch.object(adding, 'find_qq_windows', return_value=[]), \
             patch.object(adding, 'read_qq_account_from_avatar') as read_account, \
             patch.object(adding, 'read_number_rows') as read_numbers, \
             redirect_stdout(io.StringIO()):
            with self.assertRaises(RuntimeError) as caught:
                adding.main()
        read_account.assert_not_called()
        read_numbers.assert_not_called()
        return str(caught.exception)

    def test_missing_selected_process_is_actionable_error(self):
        message = self.missing_window_error(554433)
        self.assertIn('554433', message)
        self.assertIn('QQ', message)
        self.assertRegex(message, '选择|重新|登录')

    def test_no_window_in_auto_mode_is_actionable_error(self):
        message = self.missing_window_error(0)
        self.assertIn('QQ', message)
        self.assertRegex(message, '登录|主窗口')


if __name__ == '__main__':
    unittest.main()
