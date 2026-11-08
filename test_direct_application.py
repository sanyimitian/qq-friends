"""No-message applications use mocked controls; never connect to QQ."""
import unittest
from unittest.mock import patch

import open_qq_add_friend as adding
import qq_question_answer as answers
import qq_verification as original
from test_qq_verification import Control, Desktop, unconfigured_application, MESSAGE
from test_question_answer import question_form
from test_add_sequence import VirtualClock
from test_add_workflow import WorkflowHarness, fake_window

NUMBER = '851984709'


class DirectApplicationTests(unittest.TestCase):
    def prepare(self, window, answer='', timeout=0):
        return answers.prepare_verification_form(Desktop(window), 123, MESSAGE, timeout,
                                                 expected_number=NUMBER, question_answer=answer)

    def send(self, prepared):
        return answers.click_verified_send(prepared, 123, MESSAGE, expected_number=NUMBER, timeout=0)

    def test_direct_page_waits_for_stable_shape_before_preparing(self):
        window, remark, group, send, cancel = unconfigured_application()
        clock = VirtualClock()
        with patch.object(original, 'time', clock):
            prepared = self.prepare(window, timeout=5)
        self.assertGreaterEqual(clock.now, 2)
        self.assertTrue(self.send(prepared))
        self.assertEqual(remark.writes + group.writes, [])
        self.assertEqual(send.invoke_count, 1)
        self.assertEqual(cancel.invoke_count, 0)

    def test_disabled_send_with_delayed_questions_gets_uniform_answer(self):
        window, _, _, send, _ = unconfigured_application()
        send.enabled = False
        added = []
        def loaded(clock):
            if clock.sleeps == 5:
                template, edits, *_ = question_form(('你是谁？', '来做什么？'))
                window._children = template._children
                for child in window._children:
                    child._parent = window
                added.extend(edits)
        clock = VirtualClock(loaded)
        with patch.object(original, 'time', clock):
            prepared = self.prepare(window, answer='统一回答', timeout=5)
        self.assertEqual(len(prepared.answers), 2)
        self.assertEqual([edit.value for edit in added], ['统一回答', '统一回答'])
        self.assertTrue(self.send(prepared))

    def test_question_appearing_during_stability_wait_is_not_bypassed(self):
        window, *_ = unconfigured_application()
        def loaded(clock):
            if clock.sleeps == 8:
                template, *_ = question_form()
                window._children = template._children
                for child in window._children:
                    child._parent = window
        with patch.object(original, 'time', VirtualClock(loaded)):
            with self.assertRaises(answers.QuestionsNeedAnswer):
                self.prepare(window, timeout=5)

    def test_permanently_disabled_send_is_not_submitted(self):
        window, _, _, send, _ = unconfigured_application()
        send.enabled = False
        with patch.object(original, 'time', VirtualClock()):
            with self.assertRaisesRegex(RuntimeError, '发送按钮不可用'):
                self.prepare(window, timeout=3)
        self.assertEqual(send.invoke_count, 0)

    def test_new_question_after_prepare_prevents_send(self):
        window, _, _, send, _ = unconfigured_application()
        prepared = self.prepare(window)
        edit = Control('Edit', '', 'q-input__inner')
        group = Control('Group', children=[Control('Text', '你是谁？'), edit])
        group._parent = window
        window._children.insert(2, group)
        with self.assertRaisesRegex(RuntimeError, '验证问题已变化'):
            self.send(prepared)
        self.assertEqual(send.invoke_count, 0)

    def test_privacy_rejection_does_not_send(self):
        window, _, _, send, _ = unconfigured_application()
        window._children.append(Control('Text', '对方进行了隐私设置，无法添加对方为好友'))
        with self.assertRaisesRegex(RuntimeError, '隐私设置'):
            self.prepare(window)
        self.assertEqual(send.invoke_count, 0)

    def test_changed_target_blocks_direct_send(self):
        window, _, _, send, _ = unconfigured_application()
        prepared = self.prepare(window)
        window._children[1].name = '123456789'
        with self.assertRaisesRegex(RuntimeError, '目标号码不匹配'):
            self.send(prepared)
        self.assertEqual(send.invoke_count, 0)

    def test_full_workflow_sends_and_marks_direct_applications(self):
        for enabled in [True, False]:
            with self.subTest(auto_send=enabled), WorkflowHarness([fake_window(441)], target=441) as harness:
                forms = []
                def prepare(desktop, process_id, message, timeout, **kwargs):
                    window, remark, group, send, cancel = unconfigured_application(pid=process_id, number=kwargs['expected_number'])
                    forms.append((remark, group, send, cancel))
                    return answers.prepare_verification_form(Desktop(window), process_id, message, 0, **kwargs)
                harness.prepare_verification_form.side_effect = prepare
                harness.search_number.side_effect = lambda *args: ('results', 1)
                harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
                harness.click_send_button.side_effect = lambda desktop, process_id, timeout, form, expected_number: answers.click_verified_send(
                    form, process_id, adding.VERIFICATION_MESSAGE, expected_number=expected_number, timeout=0)
                with patch.object(adding, 'AUTO_SEND', enabled):
                    harness.run()
                self.assertEqual(harness.mark_rows_processed.call_count, 3 if enabled else 0)
                for remark, group, send, cancel in forms:
                    self.assertEqual(remark.writes + group.writes, [])
                    self.assertEqual(send.invoke_count, 1 if enabled else 0)
                    self.assertEqual(cancel.invoke_count, 0)


if __name__ == '__main__':
    unittest.main()
