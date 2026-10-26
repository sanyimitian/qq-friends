"""Offline navigation regressions based on the observed QQ main-window tree.

These controls are synthetic; tests never click a real QQ window or send requests.
"""

import io
import unittest
from contextlib import redirect_stdout
from dataclasses import dataclass
from types import SimpleNamespace
from unittest.mock import Mock, call, patch

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
    def __init__(self, name='', kind='Group', children=(), *,
                 rect=(0, 0, 100, 40), class_name='', process=12824,
                 visible=True, enabled=True, handle=None):
        self.name = name
        self.element_info = SimpleNamespace(
            control_type=kind, name=name, process_id=process,
            automation_id='', class_name=class_name,
        )
        self._children = list(children)
        self._parent = None
        for child in self._children:
            child._parent = self
        self._rect = Rectangle(*rect)
        self._visible = visible
        self._enabled = enabled
        self._process = process
        self.handle = id(self) if handle is None else handle
        self.click_input = Mock(name=f'click {name}')
        self.set_focus = Mock(name=f'focus {name}')

    def window_text(self):
        return self.name

    def class_name(self):
        return self.element_info.class_name

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
        return self._rect

    def is_visible(self):
        return self._visible

    def is_enabled(self):
        return self._enabled

    def process_id(self):
        return self._process


def main_window(*controls, process=12824, handle=None):
    return Control('QQ', 'Window', controls, rect=(0, 0, 2160, 1400),
                   process=process, handle=handle)


def search_edit(**kwargs):
    return Control('搜索', 'Edit', rect=(126, 75, 347, 103), **kwargs)


def header_more(class_name='contact-adder-btn top-bar__adder', **kwargs):
    return Control('更多', 'Button', rect=(362, 71, 399, 107),
                   class_name=class_name, **kwargs)


def observed_main():
    sidebar_more = Control('更多', 'Button', rect=(16, 552, 65, 602),
                           class_name='nav-item nav-item-more')
    sidebar = Control('', 'ToolBar', [sidebar_more], rect=(0, 50, 80, 1400))
    settings = Control('更多', 'Button', rect=(17, 1313, 64, 1360),
                       class_name='func-menu-more-component func-menu__item')
    settings_icon = Control('更多', 'Button', rect=(17, 1313, 64, 1360),
                            class_name='icon-item sidebar-icon')
    correct = header_more()
    header = Control('', 'Group', [search_edit(), correct], rect=(100, 55, 410, 120))
    chat_more = Control('更多', 'Button', rect=(2080, 72, 2120, 110),
                        class_name='icon-item')
    chat_toolbar = Control('更多', 'ToolBar', [chat_more],
                           rect=(1842, 69, 2141, 114),
                           class_name='func-bar-native func-bar-top')
    window = main_window(sidebar, settings, settings_icon, header, chat_toolbar)
    return window, correct, [sidebar_more, settings, settings_icon, chat_toolbar, chat_more]


class MoreEntryTests(unittest.TestCase):
    def find(self, window, extras=()):
        desktop = SimpleNamespace(windows=lambda: [*extras, window])
        return adding.find_more_menu_item(desktop, window.process_id(), window)

    def test_observed_six_more_controls_choose_top_search_adder(self):
        window, expected, rejected = observed_main()
        self.assertIs(self.find(window), expected)
        for control in rejected:
            control.click_input.assert_not_called()

    def test_explicit_main_handle_excludes_same_process_secondary_controls(self):
        target, expected, _ = observed_main()
        other_more = header_more()
        other_main = main_window(other_more, search_edit(), process=target.process_id())
        self.assertIs(self.find(target, [other_main]), expected)

    def test_other_process_header_is_not_used(self):
        target = main_window(search_edit())
        other = main_window(header_more(process=999), search_edit(process=999), process=999)
        self.assertIsNone(self.find(target, [other]))

    def test_explicit_window_process_mismatch_is_rejected(self):
        target = main_window(header_more(), search_edit())
        desktop = SimpleNamespace(windows=lambda: [target])
        self.assertIsNone(adding.find_more_menu_item(desktop, 999, target))

    def test_implicit_lookup_uses_main_qq_window_only(self):
        expected = header_more()
        main = main_window(search_edit(), expected)
        profile = Control('个人资料', 'Window', [header_more()], process=main.process_id())
        desktop = SimpleNamespace(windows=lambda: [profile, main])
        self.assertIs(adding.find_more_menu_item(desktop, main.process_id()), expected)

    def test_hidden_and_disabled_header_entries_are_not_candidates(self):
        for attr in ('visible', 'enabled'):
            with self.subTest(attr=attr):
                window = main_window(search_edit(), header_more(**{attr: False}))
                self.assertIsNone(self.find(window))

    def test_unknown_class_can_use_unique_more_beside_search_edit(self):
        expected = header_more(class_name='version-changed-header-class')
        self.assertIs(self.find(main_window(search_edit(), expected)), expected)

    def test_geometric_fallback_does_not_choose_far_right_chat_more(self):
        chat = Control('更多', 'Button', rect=(2080, 72, 2120, 110), class_name='icon-item')
        self.assertIsNone(self.find(main_window(search_edit(), chat)))

    def test_geometric_fallback_does_not_choose_bottom_settings_more(self):
        settings = Control('更多', 'Button', rect=(362, 1300, 399, 1340))
        self.assertIsNone(self.find(main_window(search_edit(), settings)))

    def test_toolbar_named_more_is_not_itself_clicked(self):
        toolbar = Control('更多', 'ToolBar', rect=(362, 71, 399, 107))
        self.assertIsNone(self.find(main_window(search_edit(), toolbar)))

    def test_two_class_matched_header_controls_are_ambiguous(self):
        self.assertIsNone(self.find(main_window(search_edit(), header_more(), header_more())))

    def test_two_geometry_matched_controls_are_ambiguous(self):
        self.assertIsNone(self.find(main_window(search_edit(),
                                               header_more(class_name=''),
                                               header_more(class_name=''))))

    def test_no_search_edit_and_unknown_more_does_not_guess(self):
        self.assertIsNone(self.find(main_window(header_more(class_name='unknown'))))


class MenuAndReuseTests(unittest.TestCase):
    def test_menu_item_lookup_keeps_target_process(self):
        wrong = Control(adding.MENU_ITEM_TEXT, 'MenuItem', process=999)
        correct = Control(adding.MENU_ITEM_TEXT, 'MenuItem')
        foreign = Control('', 'Window', [wrong], process=999)
        own = Control('', 'Window', [correct])
        desktop = SimpleNamespace(windows=lambda: [foreign, own])
        self.assertIs(adding.find_menu_item(desktop, own.process_id()), correct)

    def test_existing_page_is_reused_without_clicking_more(self):
        main, more, rejected = observed_main()
        page = Control(adding.PAGE_TITLE, 'Window')
        foreign_page = Control(adding.PAGE_TITLE, 'Window', process=999)
        desktop = SimpleNamespace(windows=lambda: [foreign_page, main, page])
        with patch.object(adding, 'activate_window'), redirect_stdout(io.StringIO()):
            self.assertEqual(adding.open_add_friend_page(desktop, Mock(), Mock(), main),
                             (main, page))
        page.set_focus.assert_called_once()
        foreign_page.set_focus.assert_not_called()
        more.click_input.assert_not_called()
        for control in rejected:
            control.click_input.assert_not_called()

    def test_already_open_add_menu_is_not_toggled_closed(self):
        main, more, _ = observed_main()
        menu_item = Control(adding.MENU_ITEM_TEXT, 'MenuItem')
        menu = Control('', 'Window', [menu_item])
        page = Control(adding.PAGE_TITLE, 'Window')
        state = {'opened': False}
        menu_item.click_input.side_effect = lambda: state.update(opened=True)
        desktop = SimpleNamespace(windows=lambda: [main, menu] + ([page] if state['opened'] else []))
        with patch.object(adding, 'activate_window'), patch.object(adding.time, 'sleep'), \
             redirect_stdout(io.StringIO()):
            result = adding.open_add_friend_page(desktop, Mock(), Mock(), main)
        self.assertEqual(result, (main, page))
        more.click_input.assert_not_called()
        menu_item.click_input.assert_called_once()

    def test_closed_menu_uses_correct_header_then_its_add_item(self):
        main, more, rejected = observed_main()
        menu_item = Control(adding.MENU_ITEM_TEXT, 'MenuItem')
        menu = Control('', 'Window', [menu_item])
        page = Control(adding.PAGE_TITLE, 'Window')
        state = {'menu': False, 'page': False}
        more.click_input.side_effect = lambda: state.update(menu=True)
        menu_item.click_input.side_effect = lambda: state.update(page=True)
        desktop = SimpleNamespace(windows=lambda: [main]
            + ([menu] if state['menu'] else []) + ([page] if state['page'] else []))
        with patch.object(adding, 'activate_window'), patch.object(adding.time, 'sleep'), \
             redirect_stdout(io.StringIO()):
            result = adding.open_add_friend_page(desktop, Mock(), Mock(), main)
        self.assertEqual(result, (main, page))
        more.click_input.assert_called_once()
        menu_item.click_input.assert_called_once()
        for control in rejected:
            control.click_input.assert_not_called()

    def test_missing_safe_entry_stops_without_clicking_bottom_more(self):
        settings = Control('更多', 'Button', rect=(17, 1313, 64, 1360))
        main = main_window(settings)
        desktop = SimpleNamespace(windows=lambda: [main])
        with patch.object(adding, 'activate_window'), self.assertRaises(RuntimeError):
            adding.open_add_friend_page(desktop, Mock(), Mock(), main)
        settings.click_input.assert_not_called()

    def test_secondary_cleanup_preserves_existing_search_page_for_reuse(self):
        main, more, _ = observed_main()
        page = Control(adding.PAGE_TITLE, 'Window')
        profile = Control('个人资料', 'Window')
        foreign = Control('个人资料', 'Window', process=999)
        all_windows = [main, page, profile, foreign]
        desktop = SimpleNamespace(windows=lambda: list(all_windows))
        gui = Mock()
        gui.IsWindowVisible.return_value = True
        gui.PostMessage.side_effect = lambda handle, *args: all_windows.__setitem__(
            slice(None), [window for window in all_windows if window.handle != handle])
        con = SimpleNamespace(WM_CLOSE=16)
        with patch.object(adding.time, 'sleep'), patch.object(adding, 'activate_window'), \
             redirect_stdout(io.StringIO()):
            adding.close_qq_secondary_windows(desktop, [main], con, gui)
            result = adding.open_add_friend_page(desktop, con, gui, main)
        self.assertEqual(result, (main, page))
        self.assertEqual(gui.PostMessage.call_args_list, [call(profile.handle, con.WM_CLOSE, 0, 0)])
        more.click_input.assert_not_called()


class SearchPageAliasTests(unittest.TestCase):
    def desktop(self, *windows):
        return SimpleNamespace(windows=lambda: list(windows))

    def test_actual_comprehensive_search_title_is_recognized(self):
        page = Control('综合搜索', 'Window')
        self.assertIs(adding.find_page_window(self.desktop(page), page.process_id()), page)

    def test_legacy_and_custom_titles_stay_supported(self):
        with patch.object(adding, 'PAGE_TITLE', '自定义搜索页'):
            for title in ('全网搜索', '综合搜索', '自定义搜索页'):
                with self.subTest(title=title):
                    page = Control(title, 'Window')
                    self.assertIs(adding.find_page_window(self.desktop(page), page.process_id()), page)

    def test_alias_lookup_ignores_other_pid_and_hidden_pages(self):
        foreign = Control('综合搜索', 'Window', process=999)
        hidden = Control('综合搜索', 'Window', visible=False)
        expected = Control('综合搜索', 'Window')
        self.assertIs(adding.find_page_window(
            self.desktop(foreign, hidden, expected), expected.process_id()), expected)
        self.assertIsNone(adding.find_page_window(self.desktop(foreign, hidden), expected.process_id()))

    def test_unrecognized_window_title_is_not_search_page(self):
        other = Control('搜索设置', 'Window')
        self.assertIsNone(adding.find_page_window(self.desktop(other), other.process_id()))

    def test_secondary_cleanup_preserves_both_builtin_aliases_and_custom_title(self):
        main = main_window()
        pages = [Control(title, 'Window') for title in ('全网搜索', '综合搜索', '自定义搜索页')]
        profile = Control('个人资料', 'Window')
        gui = Mock()
        gui.IsWindowVisible.return_value = True
        con = SimpleNamespace(WM_CLOSE=16)
        with patch.object(adding, 'PAGE_TITLE', '自定义搜索页'), \
             patch.object(adding.time, 'sleep'), redirect_stdout(io.StringIO()):
            adding.close_qq_secondary_windows(self.desktop(main, *pages, profile), [main], con, gui)
        self.assertEqual(gui.PostMessage.call_args_list, [call(profile.handle, con.WM_CLOSE, 0, 0)])

    def test_open_add_friend_reuses_actual_comprehensive_search_page(self):
        main, more, rejected = observed_main()
        page = Control('综合搜索', 'Window')
        with patch.object(adding, 'activate_window'), redirect_stdout(io.StringIO()):
            result = adding.open_add_friend_page(self.desktop(main, page), Mock(), Mock(), main)
        self.assertEqual(result, (main, page))
        page.set_focus.assert_called_once()
        more.click_input.assert_not_called()
        for control in rejected:
            control.click_input.assert_not_called()

    def test_page_titles_are_static_when_saved_static_text_configuration_is_old(self):
        with patch.object(adding, 'STATIC_PAGE_TEXTS', {'搜索', '全部', '用户'}), \
             patch.object(adding, 'PAGE_TITLE', '自定义搜索页'):
            for title in ('全网搜索', '综合搜索', '自定义搜索页'):
                with self.subTest(title=title):
                    page = Control(title, 'Window', [Control(title, 'Text')])
                    self.assertFalse(adding.has_result_content(page))

    def test_realistic_empty_document_and_actual_page_controls_resolve(self):
        search_input = Control('输入搜索关键词', 'Edit')
        search_button = Control('搜索', 'Button')
        user_tab = Control('用户', 'TabItem', class_name='q-tabs-item')
        document = Control('', 'Document', [search_input, search_button,
            Control('全部', 'TabItem', class_name='q-tabs-item active'), user_tab,
            *[Control(title, 'TabItem', class_name='q-tabs-item')
              for title in ('群聊', '频道', '小程序', '表情', '机器人')]])
        page = Control('综合搜索', 'Window', [document])
        self.assertEqual(adding.get_search_controls(page), (search_input, search_button))
        with patch.object(adding.time, 'sleep'):
            self.assertTrue(adding.select_user_tab(page))
        user_tab.click_input.assert_called_once()
        text = adding.read_search_page_text(page)
        self.assertIn('输入搜索关键词', text)
        self.assertIn('用户', text)

    def test_disabled_user_tab_is_not_clicked(self):
        tab = Control('用户', 'TabItem', class_name='q-tabs-item', enabled=False)
        page = Control('综合搜索', 'Window', [Control('', 'Document', [tab])])
        self.assertFalse(adding.select_user_tab(page))
        tab.click_input.assert_not_called()


class SearchControlTests(unittest.TestCase):
    def page(self, *controls):
        return Control(adding.PAGE_TITLE, 'Window', controls)

    def test_blank_document_aggregates_updated_visible_children(self):
        result = Control('原始结果', 'Text')
        page = self.page(Control('', 'Document', [result]))
        self.assertIn('原始结果', adding.read_search_page_text(page))
        result.name = '新的搜索结果'
        text = adding.read_search_page_text(page)
        self.assertIn('新的搜索结果', text)
        self.assertNotIn('原始结果', text)

    def test_static_document_name_does_not_mask_updated_child_text(self):
        result = Control('原始结果', 'Text')
        page = self.page(Control('全网搜索', 'Document', [result]))
        original = adding.read_search_page_text(page)
        result.name = '新的搜索结果'
        updated = adding.read_search_page_text(page)
        self.assertNotEqual(original, updated)
        self.assertIn('新的搜索结果', updated)

    def test_hidden_no_result_text_is_not_in_aggregate(self):
        page = self.page(Control('全网搜索', 'Document', [
            Control('没有相关用户', 'Text', visible=False),
            Control('可见用户', 'Text'),
        ]))
        text = adding.read_search_page_text(page)
        self.assertNotIn('没有相关用户', text)
        self.assertIn('可见用户', text)

    def test_page_without_document_still_aggregates_visible_controls(self):
        page = self.page(Control('搜索', 'Button'), Control('可见结果', 'Text'),
                         Control('隐藏结果', 'Text', visible=False))
        text = adding.read_search_page_text(page)
        self.assertIn('搜索', text)
        self.assertIn('可见结果', text)
        self.assertNotIn('隐藏结果', text)

    def test_aggregate_deduplicates_repeated_accessible_names(self):
        page = self.page(Control('用户', 'TabItem'),
                         Control('用户', 'Text'), Control('结果', 'Text'))
        self.assertEqual(adding.read_search_page_text(page).splitlines().count('用户'), 1)

    def test_named_search_edit_preferred_to_other_usable_edits(self):
        unrelated = Control('备注', 'Edit')
        expected = Control('搜索', 'Edit')
        button = Control('搜索', 'Button')
        self.assertEqual(adding.get_search_controls(self.page(unrelated, expected, button)),
                         (expected, button))

    def test_unique_usable_unnamed_edit_is_supported(self):
        expected = Control('', 'Edit')
        button = Control('搜索', 'Button')
        self.assertEqual(adding.get_search_controls(self.page(expected, button)),
                         (expected, button))

    def test_multiple_unnamed_edits_are_ambiguous(self):
        page = self.page(Control('', 'Edit'), Control('', 'Edit'),
                         Control('搜索', 'Button'))
        with self.assertRaises(RuntimeError):
            adding.get_search_controls(page)

    def test_hidden_or_disabled_input_is_not_usable(self):
        for attr in ('visible', 'enabled'):
            with self.subTest(attr=attr):
                page = self.page(Control('搜索', 'Edit', **{attr: False}),
                                 Control('搜索', 'Button'))
                with self.assertRaises(RuntimeError):
                    adding.get_search_controls(page)

    def test_disabled_named_input_does_not_hide_usable_unique_edit(self):
        disabled = Control('搜索', 'Edit', enabled=False)
        expected = Control('', 'Edit')
        button = Control('搜索', 'Button')
        self.assertEqual(adding.get_search_controls(self.page(disabled, expected, button)),
                         (expected, button))

    def test_missing_hidden_or_disabled_search_button_stops(self):
        for button in (None, Control('搜索', 'Button', visible=False),
                       Control('搜索', 'Button', enabled=False)):
            with self.subTest(button=button):
                controls = [Control('搜索', 'Edit')]
                if button is not None:
                    controls.append(button)
                with self.assertRaises(RuntimeError):
                    adding.get_search_controls(self.page(*controls))


if __name__ == '__main__':
    unittest.main()
