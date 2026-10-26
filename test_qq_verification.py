import types
import unittest
from unittest.mock import patch

import qq_verification as verification
from task_control import StopRequested


MESSAGE = '您好，0.18稳定不降智中转免费送额度，欢迎来试，地址看头像'


class Control:
    def __init__(self, kind, name='', class_name='', children=(), pid=123):
        self.element_info = types.SimpleNamespace(control_type=kind, class_name=class_name)
        self.name = name
        self.value = name
        self._children = list(children)
        self._parent = None
        self.pid = pid
        self.visible = self.enabled = True
        self.clicks = 0
        self.writes = []
        self.set_error = None
        self.get_error = None
        self.truncate = None
        self.read_values = []
        self.invoke_count = 0
        self.invoke_attempts = 0
        self.invoke_error = None
        self.invoke_close_after = 0
        self.pending_close = None
        self.visible_error = None
        self.handle = 0
        self.focus_count = 0
        self.invoke_available = True
        self.on_invoke = None
        for child in self._children:
            child._parent = self

    def descendants(self):
        return [item for child in self._children for item in [child, *child.descendants()]]

    def children(self):
        return self._children

    def parent(self):
        return self._parent

    def process_id(self):
        return self.pid

    def window_text(self):
        return self.name

    def is_visible(self):
        if self.visible_error:
            raise self.visible_error
        if self.pending_close is not None:
            if self.pending_close <= 0:
                self.visible = False
                self.pending_close = None
            else:
                self.pending_close -= 1
        return self.visible

    def is_enabled(self):
        return self.enabled

    def set_edit_text(self, value):
        if self.set_error:
            raise self.set_error
        self.writes.append(value)
        self.value = value if self.truncate is None else value[:self.truncate]

    def get_value(self):
        if self.get_error:
            raise self.get_error
        return self.read_values.pop(0) if self.read_values else self.value

    def click_input(self):
        self.clicks += 1

    def set_focus(self):
        self.focus_count += 1

    @property
    def iface_invoke(self):
        control = self

        class Invoke:
            def Invoke(self):
                control.invoke_attempts += 1
                if control.invoke_error:
                    raise control.invoke_error
                control.invoke_count += 1
                control.clicks += 1
                if control._parent is not None:
                    control._parent.pending_close = control.invoke_close_after
                    control._parent.invoked_any = True
                    if control.on_invoke:
                        control.on_invoke(control._parent)

        if not self.invoke_available:
            raise AttributeError('InvokePattern is unavailable')
        return Invoke()


class Desktop:
    def __init__(self, *windows):
        self.items = list(windows)

    def windows(self):
        return self.items


class StaleText(Control):
    def __init__(self):
        super().__init__('Text')

    def window_text(self):
        raise RuntimeError('private stale control content')


def application(pid=123):
    edit = Control('Edit', '我是原验证', 'q-textarea__inner')
    remark = Control('Edit', '原备注', 'q-input__inner')
    group = Control('Edit', '我的好友', 'q-input__inner')
    send = Control('Button', '发送')
    window = Control('Window', '加好友/群', children=[
        Control('Text', '申请加好友'),
        Control('Text', '851984709'),
        Control('Group', children=[Control('Text', '填写验证信息'), edit]),
        Control('Group', children=[Control('Text', '备注'), remark]),
        Control('Group', children=[Control('Text', '分组'), group]),
        send,
    ], pid=pid)
    send.invoke_close_after = 0
    return window, edit, remark, group, send


def unconfigured_application(pid=123, number='851984709'):
    remark = Control('Edit', '原备注', 'q-input__inner')
    group = Control('Edit', '我的好友', 'q-input__inner')
    send = Control('Button', '发送')
    cancel = Control('Button', '取消')
    window = Control('Window', '加好友/群', children=[
        Control('Text', '申请加好友'),
        Control('Text', number),
        Control('Group', children=[Control('Text', '备注'), remark]),
        Control('Group', children=[Control('Text', '分组'), group]),
        Control('Text', '好友权限'),
        send,
        cancel,
    ], pid=pid)
    return window, remark, group, send, cancel


class VerificationTests(unittest.TestCase):
    def prepare(self, desktop, message=MESSAGE, expected_number=None):
        return verification.prepare_verification_form(
            desktop, 123, message, timeout=0, expected_number=expected_number
        )

    def test_prepare_changes_only_verification_and_never_sends(self):
        window, edit, remark, group, send = application()
        self.assertIs(self.prepare(Desktop(window)), window)
        self.assertEqual(edit.value, MESSAGE)
        self.assertEqual(remark.value, '原备注')
        self.assertEqual(group.value, '我的好友')
        self.assertEqual(remark.writes + group.writes, [])
        self.assertEqual(send.clicks, 0)

    def test_stale_unrelated_text_does_not_hide_complete_form(self):
        window, edit, remark, group, send = application()
        window._children.insert(0, StaleText())
        self.assertIs(self.prepare(Desktop(window)), window)
        self.assertEqual(edit.value, MESSAGE)
        self.assertEqual(remark.writes + group.writes, [])
        self.assertEqual(send.clicks, 0)
        self.assertTrue(verification.click_verified_send(window, 123, MESSAGE))
        self.assertEqual(send.clicks, 1)

    def test_missing_verification_label_reports_form_without_editing_remark(self):
        window, edit, remark, group, send = application()
        window._children.remove(edit.parent())
        with self.assertRaisesRegex(RuntimeError, '检测到好友申请页，但未显示“填写验证信息”标签') as caught:
            self.prepare(Desktop(window))
        detail = str(caught.exception)
        self.assertIn('加好友/群', detail)
        self.assertIn('标签=申请加好友、备注、分组', detail)
        self.assertIn('Edit=q-input__inner', detail)
        self.assertIn('发送', detail)
        for private in ('原备注', '我的好友', '我是原验证'):
            self.assertNotIn(private, detail)
        self.assertEqual(edit.writes + remark.writes + group.writes, [])
        self.assertEqual(send.clicks, 0)

    def test_exact_target_without_verification_raises_special_exception(self):
        window, remark, group, send, cancel = unconfigured_application()
        with self.assertRaises(verification.NoVerificationField) as caught:
            self.prepare(Desktop(window), expected_number='851984709')
        self.assertIs(caught.exception.window, window)
        self.assertEqual(remark.writes + group.writes, [])
        self.assertEqual((send.invoke_count, send.clicks, cancel.invoke_count), (0, 0, 0))

    def test_unconfigured_form_for_wrong_number_is_not_accepted(self):
        window, remark, group, send, cancel = unconfigured_application(number='123456789')
        with self.assertRaisesRegex(RuntimeError, '目标号码不匹配'):
            self.prepare(Desktop(window), expected_number='851984709')
        self.assertEqual(remark.writes + group.writes, [])
        self.assertEqual((send.invoke_count, cancel.invoke_count), (0, 0))

    def test_verified_form_with_wrong_number_is_not_written(self):
        window, edit, remark, group, send = application()
        with self.assertRaisesRegex(RuntimeError, '目标号码不匹配'):
            self.prepare(Desktop(window), expected_number='123456789')
        self.assertEqual(edit.writes + remark.writes + group.writes, [])
        self.assertEqual(send.invoke_attempts, 0)

    def test_no_verification_form_can_be_cancelled_without_sending(self):
        window, remark, group, send, cancel = unconfigured_application()
        self.assertTrue(verification.cancel_unconfigured_form(
            window, 123, '851984709', timeout=0
        ))
        self.assertEqual(cancel.invoke_count, 1)
        self.assertEqual(send.invoke_count, 0)
        self.assertEqual(send.clicks, 0)
        self.assertEqual(remark.writes + group.writes, [])

    def test_no_verification_cancel_rejects_wrong_account_or_number(self):
        wrong_pid, _, _, pid_send, pid_cancel = unconfigured_application(pid=456)
        wrong_number, _, _, number_send, number_cancel = unconfigured_application(number='123456789')
        cases = (
            (wrong_pid, 123, '851984709', pid_send, pid_cancel),
            (wrong_number, 123, '851984709', number_send, number_cancel),
        )
        for form, process_id, number, send, cancel in cases:
            with self.subTest(process_id=process_id, number=number):
                with self.assertRaises(RuntimeError):
                    verification.cancel_unconfigured_form(form, process_id, number, timeout=0)
                self.assertEqual((cancel.invoke_count, send.invoke_count), (0, 0))

    def test_no_verification_cancel_rejects_textarea_form(self):
        window, edit, *_ = application()
        with self.assertRaisesRegex(RuntimeError, '存在验证信息区域'):
            verification.cancel_unconfigured_form(window, 123, '851984709', timeout=0)
        self.assertEqual(edit.writes, [])

    def test_no_verification_cancel_waits_for_window_to_disappear(self):
        window, _, _, send, cancel = unconfigured_application()
        cancel.invoke_close_after = 2
        with patch.object(verification.time, 'sleep'):
            self.assertTrue(verification.cancel_unconfigured_form(
                window, 123, '851984709', timeout=1
            ))
        self.assertEqual(cancel.invoke_count, 1)
        self.assertEqual(send.invoke_count, 0)

    def test_unconfigured_cancel_timeout_does_not_claim_success(self):
        window, _, _, send, cancel = unconfigured_application()
        cancel.invoke_close_after = 100
        with patch.object(verification.time, 'sleep'):
            with self.assertRaisesRegex(RuntimeError, '申请窗口尚未关闭'):
                verification.cancel_unconfigured_form(window, 123, '851984709', timeout=0)
        self.assertEqual(cancel.invoke_count, 1)
        self.assertEqual(send.invoke_count, 0)

    def test_cancel_invoke_error_never_falls_back_to_coordinate_click(self):
        window, _, _, send, cancel = unconfigured_application()
        cancel.invoke_error = RuntimeError('invoke failed')
        with self.assertRaisesRegex(RuntimeError, '不会点击发送'):
            verification.cancel_unconfigured_form(window, 123, '851984709', timeout=0)
        self.assertEqual(cancel.invoke_count, 0)
        self.assertEqual(cancel.clicks, 0)
        self.assertEqual(send.invoke_count, 0)

    def test_cancel_missing_invoke_pattern_never_falls_back(self):
        window, _, _, send, cancel = unconfigured_application()
        cancel.invoke_available = False
        with self.assertRaisesRegex(RuntimeError, '不支持安全的 UIA Invoke'):
            verification.cancel_unconfigured_form(window, 123, '851984709', timeout=0)
        self.assertEqual(cancel.clicks, 0)
        self.assertEqual(send.invoke_attempts, 0)

    def test_timeout_diagnostic_does_not_expose_other_account(self):
        other, _, _, _, _ = application(pid=456)
        other.window_text = unittest.mock.Mock(side_effect=AssertionError('must not read another account'))
        with self.assertRaisesRegex(RuntimeError, '未检测到申请页') as caught:
            self.prepare(Desktop(other))
        other.window_text.assert_not_called()
        self.assertNotIn('加好友/群', str(caught.exception))
        self.assertNotIn('q-textarea__inner', str(caught.exception))

    def test_unknown_titles_labels_buttons_and_edit_contents_are_redacted(self):
        edit = Control('Edit', 'private edit value', 'private-editor-class')
        unknown_button = Control('Button', 'private button name')
        chat = Control('Window', 'private contact nickname', children=[
            Control('Text', 'private chat text'), edit, unknown_button,
            Control('Button', '发送'), StaleText(),
        ])
        with self.assertRaisesRegex(RuntimeError, '未检测到申请页') as caught:
            self.prepare(Desktop(chat))
        detail = str(caught.exception)
        self.assertIn('其他窗口', detail)
        self.assertIn('其他Edit类', detail)
        self.assertNotIn('private', detail)
        self.assertEqual(edit.writes, [])
        self.assertEqual(unknown_button.clicks, 0)

    def test_known_window_title_without_required_labels_still_fails(self):
        edit = Control('Edit', 'private remark', 'q-input__inner')
        window = Control('Window', '加好友/群', children=[
            Control('Group', children=[Control('Text', '备注'), edit]),
            Control('Button', '发送'),
        ])
        with self.assertRaisesRegex(RuntimeError, '未识别到完整的好友申请标签'):
            self.prepare(Desktop(window))
        self.assertEqual(edit.writes, [])

    def test_stale_required_label_cannot_be_replaced_by_window_title(self):
        window, edit, _, _, send = application()
        label = edit.parent().children()[0]
        label.window_text = unittest.mock.Mock(side_effect=RuntimeError('private label error'))
        with self.assertRaisesRegex(RuntimeError, '未显示“填写验证信息”标签'):
            self.prepare(Desktop(window))
        self.assertEqual(edit.writes, [])
        self.assertEqual(send.clicks, 0)

    def test_stop_during_diagnostic_is_not_swallowed(self):
        window = Control('Window', 'QQ')
        window.window_text = unittest.mock.Mock(side_effect=StopRequested())
        with self.assertRaises(StopRequested):
            self.prepare(Desktop(window))

    def test_wrong_account_is_ignored(self):
        window, edit, _, _, send = application(pid=456)
        with self.assertRaisesRegex(RuntimeError, '未找到'):
            self.prepare(Desktop(window))
        self.assertEqual(edit.writes, [])
        self.assertEqual(send.clicks, 0)

    def test_chat_send_does_not_make_application_form(self):
        send = Control('Button', '发送')
        chat = Control('Window', 'QQ', children=[Control('Edit'), send])
        with self.assertRaisesRegex(RuntimeError, '未找到'):
            self.prepare(Desktop(chat))
        with self.assertRaisesRegex(RuntimeError, '不是完整'):
            verification.click_verified_send(chat, 123, MESSAGE)
        self.assertEqual(send.clicks, 0)

    def test_other_account_and_chat_never_receive_send_click(self):
        window, _, _, _, send = application()
        other, other_edit, _, _, other_send = application(pid=456)
        chat_send = Control('Button', '发送')
        chat = Control('Window', 'QQ', children=[chat_send])
        self.prepare(Desktop(other, chat, window))
        self.assertTrue(verification.click_verified_send(window, 123, MESSAGE))
        self.assertEqual((send.clicks, other_send.clicks, chat_send.clicks), (1, 0, 0))
        self.assertEqual(other_edit.writes, [])

    def test_multiple_forms_rejected_before_writing(self):
        first, first_edit, *_ = application()
        second, second_edit, *_ = application()
        with self.assertRaisesRegex(RuntimeError, '多个好友申请表单'):
            self.prepare(Desktop(first, second))
        self.assertEqual(first_edit.writes + second_edit.writes, [])

    def test_multiple_edits_under_verification_label_rejected(self):
        window, edit, *_ = application()
        edit.parent()._children.append(Control('Edit', 'second', 'q-textarea__inner'))
        with self.assertRaisesRegex(RuntimeError, '唯一'):
            self.prepare(Desktop(window))
        self.assertEqual(edit.writes, [])

    def test_duplicate_verification_label_rejected(self):
        window, edit, *_ = application()
        edit.parent()._children.append(Control('Text', '填写验证信息'))
        with self.assertRaisesRegex(RuntimeError, '多个验证信息区域'):
            self.prepare(Desktop(window))
        self.assertEqual(edit.writes, [])

    def test_missing_verification_edit_does_not_fall_back_to_remark(self):
        window, edit, remark, *_ = application()
        edit.parent()._children.remove(edit)
        with self.assertRaisesRegex(RuntimeError, '唯一'):
            self.prepare(Desktop(window))
        self.assertEqual(remark.writes, [])

    def test_wrong_edit_class_rejected(self):
        window, edit, *_ = application()
        edit.element_info.class_name = 'q-input__inner'
        with self.assertRaisesRegex(RuntimeError, '类型'):
            self.prepare(Desktop(window))
        self.assertEqual(edit.writes, [])

    def test_disabled_edit_rejected(self):
        window, edit, *_ = application()
        edit.enabled = False
        with self.assertRaisesRegex(RuntimeError, '不可编辑'):
            self.prepare(Desktop(window))
        self.assertEqual(edit.writes, [])

    def test_read_only_set_error_is_not_swallowed(self):
        window, edit, _, _, send = application()
        edit.set_error = RuntimeError('read only')
        with self.assertRaisesRegex(RuntimeError, '无法填写'):
            self.prepare(Desktop(window))
        self.assertEqual(send.clicks, 0)

    def test_unreadable_message_rejected(self):
        window, edit, _, _, send = application()
        edit.get_error = RuntimeError('no value pattern')
        with self.assertRaisesRegex(RuntimeError, '无法回读'):
            self.prepare(Desktop(window))
        self.assertEqual(send.clicks, 0)

    def test_truncated_message_rejected(self):
        window, edit, _, _, send = application()
        edit.truncate = 10
        with self.assertRaisesRegex(RuntimeError, '截断'):
            self.prepare(Desktop(window))
        self.assertEqual(send.clicks, 0)

    def test_crlf_is_only_normalization(self):
        window, edit, _, _, send = application()
        edit.read_values = ['第一行\r\n第二行']
        self.prepare(Desktop(window), '第一行\n第二行')
        self.assertTrue(verification.click_verified_send(window, 123, '第一行\n第二行'))
        self.assertEqual(send.clicks, 1)

    def test_send_waits_for_window_to_disappear(self):
        window, _, _, _, send = application()
        self.prepare(Desktop(window))
        send.invoke_close_after = 2
        with patch.object(verification.time, 'sleep'):
            self.assertTrue(verification.click_verified_send(
                window, 123, MESSAGE, expected_number='851984709', timeout=1
            ))
        self.assertEqual(send.invoke_attempts, 1)

    def test_send_open_window_timeout_is_uncertain_not_success(self):
        window, _, _, _, send = application()
        self.prepare(Desktop(window))
        send.invoke_close_after = 100
        with patch.object(verification.time, 'sleep'):
            with self.assertRaisesRegex(RuntimeError, '申请窗口尚未关闭'):
                verification.click_verified_send(
                    window, 123, MESSAGE, expected_number='851984709', timeout=0
                )
        self.assertEqual(send.invoke_attempts, 1)

    def test_send_invoke_exception_is_not_retried_by_click_input(self):
        window, _, _, _, send = application()
        self.prepare(Desktop(window))
        send.invoke_error = RuntimeError('invoke failed')
        with self.assertRaises(RuntimeError):
            verification.click_verified_send(window, 123, MESSAGE, expected_number='851984709')
        self.assertEqual(send.invoke_attempts, 1)
        self.assertEqual(send.clicks, 0)

    def test_send_visibility_read_error_is_not_success(self):
        window, _, _, _, send = application()
        self.prepare(Desktop(window))
        send.invoke_close_after = 100
        send.on_invoke = lambda target: setattr(target, 'visible_error', RuntimeError('transient UIA failure'))
        with self.assertRaisesRegex(RuntimeError, '无法确认申请窗口是否关闭'):
            verification.click_verified_send(window, 123, MESSAGE, timeout=1)
        self.assertEqual(send.invoke_attempts, 1)

    def test_delayed_accessible_value_is_polled(self):
        window, edit, *_ = application()
        edit.read_values = ['old', MESSAGE]
        with patch.object(verification.time, 'sleep') as sleep:
            self.assertIs(verification.prepare_verification_form(Desktop(window), 123, MESSAGE), window)
        sleep.assert_called_once_with(0.1)

    def test_modified_message_blocks_send(self):
        window, edit, _, _, send = application()
        self.prepare(Desktop(window))
        edit.value = MESSAGE + ' '
        with self.assertRaisesRegex(RuntimeError, '不一致'):
            verification.click_verified_send(window, 123, MESSAGE)
        self.assertEqual(send.clicks, 0)

    def test_message_changes_after_button_lookup_blocks_send(self):
        window, edit, _, _, send = application()
        self.prepare(Desktop(window))
        edit.read_values = [MESSAGE, 'changed']
        with self.assertRaisesRegex(RuntimeError, '不一致'):
            verification.click_verified_send(window, 123, MESSAGE)
        self.assertEqual(send.clicks, 0)

    def test_wrong_account_on_send_rejected(self):
        window, _, _, _, send = application()
        self.prepare(Desktop(window))
        window.pid = 456
        with self.assertRaisesRegex(RuntimeError, '不属于'):
            verification.click_verified_send(window, 123, MESSAGE)
        self.assertEqual(send.clicks, 0)

    def test_closed_form_rejected_on_send(self):
        window, _, _, _, send = application()
        self.prepare(Desktop(window))
        window.visible = False
        with self.assertRaisesRegex(RuntimeError, '已关闭'):
            verification.click_verified_send(window, 123, MESSAGE)
        self.assertEqual(send.clicks, 0)

    def test_missing_disabled_or_multiple_send_buttons_rejected(self):
        for mode in ('missing', 'disabled', 'multiple'):
            with self.subTest(mode=mode):
                window, _, _, _, send = application()
                self.prepare(Desktop(window))
                if mode == 'missing':
                    window._children.remove(send)
                elif mode == 'disabled':
                    send.enabled = False
                else:
                    window._children.append(Control('Button', '发送'))
                with self.assertRaisesRegex(RuntimeError, '唯一可用'):
                    verification.click_verified_send(window, 123, MESSAGE)
                self.assertEqual(send.clicks, 0)

    def test_stop_during_write_propagates(self):
        window, edit, *_ = application()
        edit.set_error = StopRequested()
        with self.assertRaises(StopRequested):
            self.prepare(Desktop(window))

    def test_stop_during_read_propagates(self):
        window, edit, *_ = application()
        edit.get_error = StopRequested()
        with self.assertRaises(StopRequested):
            self.prepare(Desktop(window))

    def test_blank_message_rejected_before_writing(self):
        window, edit, *_ = application()
        with self.assertRaisesRegex(ValueError, '不能为空'):
            self.prepare(Desktop(window), ' \n ')
        self.assertEqual(edit.writes, [])


if __name__ == '__main__':
    unittest.main()
