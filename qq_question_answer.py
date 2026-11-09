"""Prepare question answers or a verified application without a message field."""
from dataclasses import dataclass

import qq_verification as original
from task_control import check_stop


@dataclass(frozen=True)
class PreparedQuestions:
    window: object
    process_id: int
    number: str
    answers: tuple


class QuestionsNeedAnswer(original.NoVerificationField):
    def __init__(self, window, questions):
        super().__init__(window)
        self.reason = f'对方设置了 {len(questions)} 个验证问题；未配置统一回答，已取消且未发送'
        self.args = (self.reason,)


def _questions(window, process_id, number):
    original._validate_expected_number(number)
    if number is None:
        raise ValueError('回答验证问题必须提供目标 QQ 号码。')
    if window.process_id() != process_id or not window.is_visible():
        raise RuntimeError('好友申请表单已关闭或不属于所选 QQ，已停止操作。')
    texts = original._visible_texts(window)
    if window.window_text().strip() not in original._APPLICATION_TITLES or original._FORM_LABEL not in texts:
        raise RuntimeError('当前窗口不是好友申请表单，已停止操作。')
    if number not in texts:
        raise RuntimeError('好友申请表单目标号码不匹配，已停止操作。')
    if '对方进行了隐私设置，无法添加对方为好友' in texts:
        raise RuntimeError('对方进行了隐私设置，无法添加对方为好友；未发送。')
    if original._MESSAGE_LABEL in texts or original._has_verification_textarea(window):
        raise RuntimeError('好友申请的验证方式已变化，已停止操作。')
    questions, auxiliary = [], []
    for control in window.descendants():
        if control.element_info.control_type != 'Edit' or not control.is_visible():
            continue
        if 'q-input__inner' not in control.element_info.class_name.split():
            raise RuntimeError('申请页出现无法识别的输入框，未填写或发送。')
        siblings = [c for c in control.parent().children() if c.is_visible()]
        labels = [c.window_text().strip() for c in siblings
                  if c.element_info.control_type == 'Text' and c.window_text().strip()]
        edits = [c for c in siblings if c.element_info.control_type == 'Edit']
        if len(labels) != 1 or len(edits) != 1:
            raise RuntimeError('申请页输入框与标签对应不唯一，未填写或发送。')
        if labels[0] in {'备注', '分组'}:
            auxiliary.append(labels[0])
        else:
            questions.append((labels[0], control))
    if sorted(auxiliary) != ['分组', '备注'] or '好友权限' not in texts:
        raise RuntimeError('申请页的备注、分组或权限尚未加载完整，未发送。')
    if len({question for question, _ in questions}) != len(questions):
        raise RuntimeError('申请页出现重复验证问题，无法唯一对应答案。')
    return tuple(questions)


def _validate_answers(prepared):
    questions = _questions(prepared.window, prepared.process_id, prepared.number)
    if tuple(q for q, _ in questions) != tuple(q for q, _ in prepared.answers):
        raise RuntimeError('好友申请的验证问题已变化，已停止发送。')
    for (question, edit), (_, answer) in zip(questions, prepared.answers):
        if not edit.is_enabled() or original._normalized(edit.get_value()) != original._normalized(answer):
            raise RuntimeError(f'验证问题“{question}”的答案回读不一致，已停止发送。')


def prepare_verification_form(desktop, process_id, message, timeout=5, *,
                              expected_number=None, question_answer=''):
    if not isinstance(question_answer, str):
        raise ValueError('验证问题统一回答必须是文本。')
    deadline = original.time.monotonic() + max(0, timeout)
    try:
        return original.prepare_verification_form(
            desktop, process_id, message, timeout, expected_number=expected_number)
    except original.NoVerificationField as missing:
        # Reacquire the exact target: UIA can briefly omit a refreshing window.
        while True:
            check_stop()
            forms = {}
            for candidate in desktop.windows():
                try:
                    if (candidate.process_id() != process_id or not candidate.is_visible()
                            or candidate.window_text().strip() not in original._APPLICATION_TITLES):
                        continue
                    texts = original._visible_texts(candidate)
                    if original._FORM_LABEL not in texts or expected_number not in texts:
                        continue
                    handle = getattr(candidate, 'handle', None)
                    key = ('handle', handle) if type(handle) is int and handle > 0 else ('object', id(candidate))
                    forms[key] = candidate
                except Exception:
                    continue
            if len(forms) > 1:
                raise RuntimeError(f'目标 QQ 有多个好友申请表单（{len(forms)} 个），未填写或发送。')
            if forms:
                window = next(iter(forms.values()))
                break
            if original.time.monotonic() >= deadline:
                raise RuntimeError('目标 QQ 的好友申请表单未加载或已关闭（找到 0 个），未填写或发送。')
            original.time.sleep(0.1)
        if original._MESSAGE_LABEL in original._visible_texts(window):
            return original.prepare_verification_form(
                desktop, process_id, message, max(0, deadline - original.time.monotonic()),
                expected_number=expected_number)
        questions = _questions(window, process_id, expected_number)
        signature = tuple(q for q, _ in questions)
        stable_since = original.time.monotonic()
        while True:
            check_stop()
            if original._MESSAGE_LABEL in original._visible_texts(window):
                return original.prepare_verification_form(
                    desktop, process_id, message, max(0, deadline - original.time.monotonic()),
                    expected_number=expected_number)
            questions = _questions(window, process_id, expected_number)
            current = tuple(q for q, _ in questions)
            now = original.time.monotonic()
            if current != signature:
                signature, stable_since = current, now
            buttons = [c for c in window.descendants()
                       if c.element_info.control_type == 'Button'
                       and c.window_text().strip() == '发送' and c.is_visible()]
            ready = bool(questions) or (len(buttons) == 1 and buttons[0].is_enabled())
            if not ready:
                stable_since = now
            stable_seconds = 1.0 if questions else 2.0
            if ready and (timeout <= 0 or now - stable_since >= stable_seconds):
                break
            if now >= deadline:
                raise RuntimeError('申请页尚未稳定加载或发送按钮不可用，未填写或发送。')
            original.time.sleep(0.1)
        if not questions:
            print('申请页没有验证信息框或验证问题，无需输入内容，已核对目标号码。', flush=True)
            return PreparedQuestions(window, process_id, expected_number, ())
        if not question_answer.strip():
            raise QuestionsNeedAnswer(window, questions) from missing
        if any(not edit.is_enabled() for _, edit in questions):
            raise RuntimeError('验证问题当前不可填写，未发送。')
        answers = tuple((question, question_answer) for question, _ in questions)
        for question, edit in questions:
            check_stop()
            current = _questions(window, process_id, expected_number)
            if tuple(q for q, _ in current) != tuple(q for q, _ in answers):
                raise RuntimeError('好友申请的验证问题已变化，已停止填写。')
            edit.set_edit_text(question_answer)
        prepared = PreparedQuestions(window, process_id, expected_number, answers)
        read_deadline = original.time.monotonic() + min(1.0, max(0, timeout))
        while True:
            check_stop()
            try:
                _validate_answers(prepared)
                return prepared
            except RuntimeError:
                if original.time.monotonic() >= read_deadline:
                    raise
                original.time.sleep(0.1)


def click_verified_send(form, process_id, message, send_button_text='发送', *,
                        expected_number=None, timeout=5):
    if not isinstance(form, PreparedQuestions):
        return original.click_verified_send(form, process_id, message, send_button_text,
                                            expected_number=expected_number, timeout=timeout)
    if form.process_id != process_id or form.number != expected_number:
        raise RuntimeError('准备好的申请与所选 QQ 或目标号码不一致，已停止发送。')
    return original._send_once(
        form.window, process_id, lambda: _validate_answers(form),
        send_button_text, expected_number, timeout)
