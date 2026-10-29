"""Search sequencing regressions using synthetic controls and a virtual clock."""
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, call, patch

import open_qq_add_friend as adding
from test_add_navigation import Control


OLD_NUMBER = '123456789'
NEW_NUMBER = '223456789'
NEXT_NUMBER = '323456789'


class SearchInput(Control):
    def __init__(self, value='', name='输入搜索关键词'):
        super().__init__(name, 'Edit', rect=(100, 60, 380, 100))
        self.value = value
        self.get_value = Mock(side_effect=lambda: self.value)
        self.set_edit_text = Mock(side_effect=self._set_value)

    def _set_value(self, value):
        self.value = value


def button(name, rect=(100, 180, 180, 220), **kwargs):
    control = Control(name, 'Button', rect=rect, **kwargs)
    control.iface_invoke = SimpleNamespace(Invoke=Mock(name=f'Invoke {name}'))
    return control


def search_page(result_number=OLD_NUMBER, *, query=OLD_NUMBER, extra=()):
    search_input = SearchInput(query)
    search_button = button('搜索', rect=(400, 60, 460, 100))
    result = Control(result_number, 'Text', rect=(100, 140, 260, 165))
    add_button = button('添加')
    document = Control('', 'Document', [search_input, search_button,
                       Control('用户', 'TabItem', rect=(100, 110, 160, 135)),
                       result, add_button, *extra])
    page = Control('综合搜索', 'Window', [document], rect=(80, 30, 700, 600))
    return page, search_input, search_button, result, add_button


class VirtualClock:
    def __init__(self, on_sleep=None):
        self.now = 0
        self.sleeps = 0
        self.on_sleep = on_sleep
        self.monotonic = Mock(side_effect=lambda: self.now)
        self.sleep = Mock(side_effect=self._sleep)

    def _sleep(self, seconds):
        self.now += max(seconds, 0.01)
        self.sleeps += 1
        if self.on_sleep:
            self.on_sleep(self)


class ResultNumberTests(unittest.TestCase):
    def test_independent_visible_number_text_matches_observed_result(self):
        page, search_input, _, _, _ = search_page(NEW_NUMBER, query=NEW_NUMBER)
        search_input.name = NEW_NUMBER
        # Accessible text aggregation deduplicates this number to one occurrence.
        self.assertEqual(adding.read_search_page_text(page).count(NEW_NUMBER), 1)
        self.assertTrue(adding.result_matches_number(page, NEW_NUMBER))

    def test_input_new_number_does_not_validate_old_result(self):
        page, search_input, _, _, _ = search_page(OLD_NUMBER, query=NEW_NUMBER)
        search_input.name = NEW_NUMBER
        self.assertFalse(adding.result_matches_number(page, NEW_NUMBER))

    def test_result_text_accepts_complete_number_with_label(self):
        page, _, _, _, _ = search_page(f'QQ号：{NEW_NUMBER}（用户）')
        self.assertTrue(adding.result_matches_number(page, NEW_NUMBER))

    def test_result_number_requires_digit_boundaries(self):
        for text in (NEW_NUMBER + '0', '0' + NEW_NUMBER, 'QQ号：' + NEW_NUMBER + '99'):
            with self.subTest(text=text):
                page, _, _, _, _ = search_page(text)
                self.assertFalse(adding.result_matches_number(page, NEW_NUMBER))

    def test_text_above_search_button_is_not_a_result(self):
        above = Control(NEW_NUMBER, 'Text', rect=(100, 35, 260, 55))
        page, _, _, _, _ = search_page(extra=[above])
        self.assertFalse(adding.result_matches_number(page, NEW_NUMBER))

    def test_hidden_result_text_is_ignored(self):
        page, _, _, result, _ = search_page(NEW_NUMBER)
        result._visible = False
        self.assertFalse(adding.result_matches_number(page, NEW_NUMBER))

    def test_non_text_controls_cannot_supply_result_identity(self):
        for kind in ('Edit', 'Group', 'Button'):
            with self.subTest(kind=kind):
                extra = Control(NEW_NUMBER, kind, rect=(100, 250, 250, 280))
                page, _, _, _, _ = search_page(extra=[extra])
                self.assertFalse(adding.result_matches_number(page, NEW_NUMBER))


class InvokeTests(unittest.TestCase):
    def test_available_invoke_activates_once_without_mouse_click(self):
        page, _, search_button, _, _ = search_page()
        adding.invoke_page_button(page, search_button)
        search_button.iface_invoke.Invoke.assert_called_once_with()
        search_button.click_input.assert_not_called()

    def test_missing_interface_focuses_page_before_mouse_fallback(self):
        page, _, _, _, _ = search_page()
        target = Control('添加', 'Button')
        events = Mock()
        events.attach_mock(page.set_focus, 'focus')
        events.attach_mock(target.click_input, 'click')
        adding.invoke_page_button(page, target)
        self.assertEqual(events.mock_calls, [call.focus(), call.click()])

    def test_unsupported_pattern_uses_single_mouse_fallback(self):
        from pywinauto.uia_defines import NoPatternInterfaceError

        class NoInvokeControl(Control):
            @property
            def iface_invoke(self):
                raise NoPatternInterfaceError('Invoke')

        page, _, _, _, _ = search_page()
        target = NoInvokeControl('添加', 'Button')
        adding.invoke_page_button(page, target)
        page.set_focus.assert_called_once_with()
        target.click_input.assert_called_once_with()

    def test_invoke_failure_never_clicks_again_even_for_attribute_error(self):
        for error in (RuntimeError('invoke failed'), AttributeError('invoke failed')):
            with self.subTest(error=type(error).__name__):
                page, _, target, _, _ = search_page()
                target.iface_invoke.Invoke.side_effect = error
                with self.assertRaises(type(error)):
                    adding.invoke_page_button(page, target)
                target.iface_invoke.Invoke.assert_called_once_with()
                target.click_input.assert_not_called()


class ResultAddButtonTests(unittest.TestCase):
    def test_unique_result_add_button_is_focused_and_invoked_once(self):
        above = button('添加', rect=(100, 35, 180, 55))
        hidden = button('添加', visible=False)
        disabled = button('添加', enabled=False)
        page, _, _, _, add_button = search_page(NEW_NUMBER, extra=[above, hidden, disabled])
        events = Mock()
        events.attach_mock(page.set_focus, 'focus')
        events.attach_mock(add_button.iface_invoke.Invoke, 'invoke')
        self.assertTrue(adding.click_result_add_button(page, expected_number=NEW_NUMBER))
        self.assertLess(events.mock_calls.index(call.focus()), events.mock_calls.index(call.invoke()))
        add_button.iface_invoke.Invoke.assert_called_once_with()
        add_button.click_input.assert_not_called()
        for rejected in (above, hidden, disabled):
            rejected.iface_invoke.Invoke.assert_not_called()
            rejected.click_input.assert_not_called()

    def test_wrong_result_number_stops_before_any_add(self):
        page, _, _, _, add_button = search_page(OLD_NUMBER, query=NEW_NUMBER)
        with self.assertRaises(RuntimeError):
            adding.click_result_add_button(page, expected_number=NEW_NUMBER)
        add_button.iface_invoke.Invoke.assert_not_called()
        add_button.click_input.assert_not_called()

    def test_multiple_visible_result_add_buttons_are_ambiguous(self):
        second = button('添加', rect=(100, 300, 180, 340))
        page, _, _, _, first = search_page(NEW_NUMBER, extra=[second])
        with self.assertRaises(RuntimeError):
            adding.click_result_add_button(page, expected_number=NEW_NUMBER)
        for target in (first, second):
            target.iface_invoke.Invoke.assert_not_called()
            target.click_input.assert_not_called()

    def test_absent_result_add_button_returns_false(self):
        page, _, _, _, target = search_page(NEW_NUMBER)
        target.name = '发消息'
        self.assertFalse(adding.click_result_add_button(page, expected_number=NEW_NUMBER))
        target.iface_invoke.Invoke.assert_not_called()


class SearchStateTests(unittest.TestCase):
    def test_changed_and_stable_old_result_never_counts_as_new_search(self):
        page, search_input, _, _, _ = search_page(OLD_NUMBER, query=NEW_NUMBER)
        search_input.name = NEW_NUMBER
        clock = VirtualClock()
        with patch.object(adding, 'time', clock), patch.object(adding, 'RESULT_STABLE_SECONDS', 0.25):
            state = adding.wait_for_search_state(page, 'old baseline', 1, True,
                                                 expected_number=NEW_NUMBER)
        self.assertEqual(state, 'timeout')

    def test_wait_accepts_correct_result_only_after_old_result_replaced(self):
        page, _, _, result, _ = search_page(OLD_NUMBER, query=NEW_NUMBER)

        def update(clock):
            if clock.sleeps == 3:
                result.name = NEW_NUMBER

        clock = VirtualClock(update)
        with patch.object(adding, 'time', clock), patch.object(adding, 'RESULT_STABLE_SECONDS', 0.25):
            state = adding.wait_for_search_state(page, 'old baseline', 3, True,
                                                 expected_number=NEW_NUMBER)
        self.assertEqual(state, 'results')
        self.assertGreaterEqual(clock.sleeps, 4)
        self.assertEqual(result.name, NEW_NUMBER)

    def test_explicit_no_results_remains_accepted_without_number_text(self):
        page, _, _, result, add_button = search_page('没有相关用户', query=NEW_NUMBER)
        add_button._visible = False
        clock = VirtualClock()
        with patch.object(adding, 'time', clock), patch.object(adding, 'RESULT_STABLE_SECONDS', 0.25):
            state = adding.wait_for_search_state(page, 'old baseline', 2, True,
                                                 expected_number=NEW_NUMBER)
        self.assertEqual(state, 'no_results')


class SearchSequenceTests(unittest.TestCase):
    def test_each_search_refocuses_types_confirms_and_invokes_for_expected_number(self):
        page, search_input, search_button, _, _ = search_page()
        events = Mock()
        events.attach_mock(page.set_focus, 'focus')
        events.attach_mock(search_input.set_edit_text, 'type')
        events.attach_mock(search_button.iface_invoke.Invoke, 'invoke')
        clock = VirtualClock()
        with patch.object(adding, 'time', clock), \
             patch.object(adding, 'wait_for_search_state', return_value='results') as wait:
            for number in (NEW_NUMBER, NEXT_NUMBER):
                events.reset_mock()
                self.assertEqual(adding.search_number(page, number, True), ('results', 1))
                calls = events.mock_calls
                self.assertLess(calls.index(call.focus()), calls.index(call.type(number)))
                self.assertLess(calls.index(call.type(number)), calls.index(call.invoke()))
                self.assertEqual(wait.call_args.kwargs.get('expected_number'), number)
                self.assertEqual(search_input.value, number)
        self.assertEqual(search_button.iface_invoke.Invoke.call_count, 1)
        # reset_mock above resets attached children each iteration; each invokes once.
        self.assertEqual(wait.call_count, 2)
        search_input.get_value.assert_called()
        search_button.click_input.assert_not_called()

    def test_failed_input_readback_stops_before_search_or_wait(self):
        page, search_input, search_button, _, _ = search_page()
        search_input.set_edit_text.side_effect = lambda value: None
        with patch.object(adding, 'time', VirtualClock()), \
             patch.object(adding, 'wait_for_search_state') as wait:
            with self.assertRaises(RuntimeError):
                adding.search_number(page, NEW_NUMBER, True)
        search_input.get_value.assert_called()
        search_button.iface_invoke.Invoke.assert_not_called()
        search_button.click_input.assert_not_called()
        wait.assert_not_called()

    def test_search_invoke_failure_does_not_retry_or_mouse_click(self):
        page, _, search_button, _, _ = search_page()
        search_button.iface_invoke.Invoke.side_effect = RuntimeError('search invoke failed')
        with patch.object(adding, 'time', VirtualClock()), \
             patch.object(adding, 'MAX_SEARCH_ATTEMPTS', 3), \
             patch.object(adding, 'wait_for_search_state') as wait:
            with self.assertRaises(RuntimeError):
                adding.search_number(page, NEW_NUMBER, True)
        search_button.iface_invoke.Invoke.assert_called_once_with()
        search_button.click_input.assert_not_called()
        wait.assert_not_called()


if __name__ == '__main__':
    unittest.main(verbosity=2)
