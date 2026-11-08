"""Reproduce the bundled wrapper race and changing QQ search trees offline."""
import io
from contextlib import ExitStack, redirect_stdout
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import open_qq_add_friend as adding
from qq_uia_compat import install_uia_wrapper_fix
from task_control import StopRequested
from test_add_sequence import search_page, VirtualClock, NEW_NUMBER, OLD_NUMBER


class UiaCompatibilityTests(unittest.TestCase):
    def setUp(self):
        from pywinauto.controls.uiawrapper import UIAWrapper, UiaMeta
        self.meta = UiaMeta
        self.wrapper = UIAWrapper
        original = UiaMeta.__dict__['find_wrapper']
        self.addCleanup(setattr, UiaMeta, 'find_wrapper', original)

    def test_control_type_that_disappears_is_read_only_once(self):
        class ChangingElement:
            reads = 0
            @property
            def control_type(self):
                self.reads += 1
                return 'Button' if self.reads == 1 else None
        element = ChangingElement()
        install_uia_wrapper_fix()
        self.assertIs(self.meta.find_wrapper(element), self.meta.control_type_to_cls['Button'])
        self.assertEqual(element.reads, 1)

    def test_none_and_unknown_types_use_generic_wrapper(self):
        install_uia_wrapper_fix()
        for kind in (None, 'UnknownQQType', 'Window'):
            self.assertIs(self.meta.find_wrapper(SimpleNamespace(control_type=kind)), self.wrapper)

    def test_known_control_types_keep_their_specific_wrappers(self):
        install_uia_wrapper_fix()
        for kind in ('Button', 'Edit', 'Text'):
            self.assertIs(self.meta.find_wrapper(SimpleNamespace(control_type=kind)), self.meta.control_type_to_cls[kind])

    def test_install_is_idempotent(self):
        install_uia_wrapper_fix()
        fixed = self.meta.find_wrapper
        install_uia_wrapper_fix()
        self.assertIs(self.meta.find_wrapper, fixed)

    def test_friend_dependency_loading_installs_compatibility_fix(self):
        with patch('qq_uia_compat.install_uia_wrapper_fix') as install:
            dependencies = adding.load_dependencies()
        install.assert_called_once_with()
        self.assertEqual(len(dependencies), 10)


class SearchRefreshTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.stack.enter_context(redirect_stdout(io.StringIO()))
        self.clock = VirtualClock()
        self.stack.enter_context(patch.object(adding, 'time', self.clock))
        self.stack.enter_context(patch.object(adding, 'RESULT_STABLE_SECONDS', 0.25))
        self.page, self.query, self.search, self.result, self.add = search_page(NEW_NUMBER, query=NEW_NUMBER)
        self.controls = self.page.descendants()

    def wait(self, timeout=3):
        return adding.wait_for_search_state(self.page, 'old baseline', timeout, True, expected_number=NEW_NUMBER)

    def test_descendant_keyerror_recovers_with_a_fresh_tree(self):
        with patch.object(self.page, 'descendants', side_effect=[KeyError(None), self.controls]) as read:
            self.assertIn(NEW_NUMBER, adding.read_search_page_text(self.page))
        self.assertEqual(read.call_count, 2)

    def test_none_control_retries_entire_tree_instead_of_skipping_it(self):
        incomplete = self.controls + [SimpleNamespace(element_info=SimpleNamespace(control_type=None))]
        with patch.object(self.page, 'descendants', side_effect=[incomplete, self.controls]) as read:
            self.assertEqual(adding.search_page_controls(self.page), self.controls)
        self.assertEqual(read.call_count, 2)

    def test_unrelated_keyerror_is_not_hidden(self):
        with patch.object(self.page, 'descendants', side_effect=KeyError('unexpected')) as read:
            with self.assertRaises(KeyError):
                adding.read_search_page_text(self.page)
        read.assert_called_once()

    def test_repeated_refresh_during_polling_eventually_recovers(self):
        calls = 0
        def descendants():
            nonlocal calls
            calls += 1
            if calls <= 4:
                raise KeyError(None)
            return self.controls
        with patch.object(self.page, 'descendants', side_effect=descendants):
            self.assertEqual(self.wait(), 'results')
        self.add.iface_invoke.Invoke.assert_not_called()

    def test_permanent_refresh_times_out_not_no_results_or_success(self):
        with patch.object(self.page, 'descendants', side_effect=KeyError(None)):
            self.assertEqual(self.wait(timeout=1), 'timeout')
        self.assertLess(self.clock.now, 1.6)
        self.add.iface_invoke.Invoke.assert_not_called()

    def test_no_results_text_in_incomplete_tree_is_not_accepted(self):
        self.result.name = '没有相关用户'
        self.add._visible = False
        incomplete = self.controls + [SimpleNamespace(element_info=SimpleNamespace(control_type=None))]
        with patch.object(self.page, 'descendants', return_value=incomplete):
            self.assertEqual(self.wait(timeout=1), 'timeout')

    def test_refresh_resets_stability_before_success(self):
        observations = ['first', 'second', adding.SearchPageNotReady('refresh'), 'stable', 'stable']
        with patch.object(adding, 'read_search_page_text', side_effect=observations):
            self.assertEqual(self.wait(), 'results')
        self.assertGreaterEqual(self.clock.sleeps, 4)

    def test_recovered_tree_must_still_match_requested_number(self):
        self.result.name = OLD_NUMBER
        self.query.name = NEW_NUMBER
        with patch.object(self.page, 'descendants', side_effect=lambda: self.controls):
            self.assertEqual(self.wait(timeout=1), 'timeout')
        self.add.iface_invoke.Invoke.assert_not_called()

    def test_add_does_not_act_on_incomplete_tree(self):
        with patch.object(self.page, 'descendants', side_effect=KeyError(None)):
            with self.assertRaises(adding.SearchPageNotReady):
                adding.click_result_add_button(self.page, expected_number=NEW_NUMBER)
        self.add.iface_invoke.Invoke.assert_not_called()
        self.add.click_input.assert_not_called()

    def test_stop_during_refresh_is_immediate(self):
        with patch.object(self.page, 'descendants', side_effect=StopRequested()) as read:
            with self.assertRaises(StopRequested):
                self.wait()
        read.assert_called_once()


if __name__ == '__main__':
    unittest.main()
