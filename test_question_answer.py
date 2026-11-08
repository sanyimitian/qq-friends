"""Offline coverage of the sole extension to the restored version."""
import unittest
from unittest.mock import patch

import open_qq_add_friend as adding
import qq_question_answer as answers
import qq_verification as original
from test_qq_verification import Control, Desktop, application, unconfigured_application, MESSAGE
from test_add_sequence import VirtualClock
from test_add_workflow import WorkflowHarness, fake_window

NUMBER = '851984709'
ANSWER = '你好，想交流一下。'


def question_form(labels=('什么事情？',)):
    window, remark, group, send, cancel = unconfigured_application()
    edits = []
    for index, label in enumerate(labels):
        edit = Control('Edit', '', 'q-input__inner')
        area = Control('Group', children=[Control('Text', label), edit])
        area._parent = window
        window._children.insert(2 + index, area)
        edits.append(edit)
    return window, edits, remark, group, send, cancel


class QuestionAnswerTests(unittest.TestCase):
    def prepare(self, window, answer=ANSWER, timeout=0):
        return answers.prepare_verification_form(Desktop(window), 123, MESSAGE, timeout,
                                                 expected_number=NUMBER, question_answer=answer)

    def send(self, form):
        return answers.click_verified_send(form, 123, MESSAGE, expected_number=NUMBER, timeout=0)

    def test_single_and_multiple_questions_use_the_same_answer(self):
        for labels in [('什么事情？',), ('你是谁？', '你多大了？')]:
            with self.subTest(labels=labels):
                window, edits, remark, group, send, _ = question_form(labels)
                form = self.prepare(window)
                self.assertEqual([edit.value for edit in edits], [ANSWER] * len(labels))
                self.assertEqual(remark.writes + group.writes, [])
                self.assertEqual(send.invoke_count, 0)
                self.assertTrue(self.send(form))
                self.assertEqual(send.invoke_count, 1)

    def test_blank_answer_uses_original_cancel_and_does_not_send(self):
        window, edits, remark, group, send, cancel = question_form(('你是谁？', '你多大了？'))
        with self.assertRaises(answers.QuestionsNeedAnswer) as caught:
            self.prepare(window, '  ')
        self.assertIn('2 个验证问题', caught.exception.reason)
        original.cancel_unconfigured_form(window, 123, NUMBER, timeout=0)
        self.assertEqual(cancel.invoke_count, 1)
        self.assertEqual(send.invoke_count, 0)
        self.assertEqual([e.writes for e in edits], [[], []])
        self.assertEqual(remark.writes + group.writes, [])

    def test_no_question_form_sends_without_writing_any_input(self):
        window, remark, group, send, cancel = unconfigured_application()
        prepared = self.prepare(window, '')
        self.assertEqual(prepared.answers, ())
        self.assertEqual(send.invoke_count, 0)
        self.assertTrue(self.send(prepared))
        self.assertEqual(send.invoke_count, 1)
        self.assertEqual(cancel.invoke_count, 0)
        self.assertEqual(remark.writes + group.writes, [])

    def test_normal_message_form_still_uses_original_functions(self):
        window, edit, remark, group, send = application()
        form = self.prepare(window)
        self.assertIs(form, window)
        self.assertEqual(edit.value, MESSAGE)
        self.assertEqual(remark.writes, [])
        self.assertTrue(self.send(form))
        self.assertEqual(send.invoke_count, 1)

    def test_changed_answer_blocks_send(self):
        window, edits, _, _, send, _ = question_form()
        form = self.prepare(window)
        edits[0].value = 'changed'
        with self.assertRaisesRegex(RuntimeError, '回读不一致'):
            self.send(form)
        self.assertEqual(send.invoke_count, 0)

    def test_truncated_answer_blocks_send(self):
        window, edits, _, _, send, _ = question_form()
        edits[0].truncate = 2
        with self.assertRaisesRegex(RuntimeError, '回读不一致'):
            self.prepare(window)
        self.assertEqual(send.invoke_count, 0)

    def test_wrong_target_and_wrong_account_block_send(self):
        window, _, _, _, send, _ = question_form()
        with self.assertRaisesRegex(RuntimeError, '目标号码不匹配'):
            answers.prepare_verification_form(Desktop(window), 123, MESSAGE, 0,
                                               expected_number='123456789', question_answer=ANSWER)
        form = self.prepare(window)
        with self.assertRaisesRegex(RuntimeError, '不一致'):
            answers.click_verified_send(form, 456, MESSAGE, expected_number=NUMBER, timeout=0)
        self.assertEqual(send.invoke_count, 0)

    def test_changed_questions_block_send(self):
        window, edits, _, _, send, _ = question_form()
        form = self.prepare(window)
        edits[0].parent().children()[0].name = '新问题'
        with self.assertRaisesRegex(RuntimeError, '验证问题已变化'):
            self.send(form)
        self.assertEqual(send.invoke_count, 0)

    def test_invoke_failure_is_not_retried(self):
        window, _, _, _, send, _ = question_form()
        form = self.prepare(window)
        send.invoke_error = RuntimeError('uncertain')
        with self.assertRaisesRegex(RuntimeError, '不会重复发送'):
            self.send(form)
        self.assertEqual(send.invoke_attempts, 1)

    def test_ambiguous_form_and_unknown_input_are_rejected(self):
        first, *_ = question_form()
        second, *_ = question_form()
        with self.assertRaisesRegex(RuntimeError, '多个好友申请'):
            answers.prepare_verification_form(Desktop(first, second), 123, MESSAGE, 0,
                                               expected_number=NUMBER, question_answer=ANSWER)
        first._children.append(Control('Edit', 'unknown', 'unknown-class'))
        with self.assertRaisesRegex(RuntimeError, '无法识别的输入框'):
            self.prepare(first)

    def test_duplicate_questions_do_not_get_partial_answers(self):
        window, edits, _, _, send, _ = question_form(('你是谁？', '你是谁？'))
        with self.assertRaisesRegex(RuntimeError, '重复验证问题'):
            self.prepare(window)
        self.assertEqual([e.writes for e in edits], [[], []])
        self.assertEqual(send.invoke_count, 0)

    def test_question_labels_stabilize_before_any_writes(self):
        window, edits, *_ = question_form()
        clock = VirtualClock()
        with patch.object(original, 'time', clock):
            self.prepare(window, timeout=3)
        self.assertGreaterEqual(clock.now, 1.0)
        self.assertEqual(edits[0].writes, [ANSWER])

    def test_main_passes_saved_answer_and_respects_original_send_switch(self):
        for auto_send in [False, True]:
            with self.subTest(auto_send=auto_send), WorkflowHarness([fake_window(441)], target=441) as harness:
                harness.search_number.side_effect = lambda *args: ('results', 1)
                harness.click_result_add_button.side_effect = lambda *args, **kwargs: True
                harness.click_send_button.side_effect = lambda *args, **kwargs: True
                with patch.object(adding, 'QUESTION_ANSWER', ANSWER), patch.object(adding, 'AUTO_SEND', auto_send):
                    harness.run()
                for call in harness.prepare_verification_form.call_args_list:
                    self.assertEqual(call.kwargs['question_answer'], ANSWER)
                self.assertEqual(harness.click_send_button.call_count, 3 if auto_send else 0)


if __name__ == '__main__':
    unittest.main()
