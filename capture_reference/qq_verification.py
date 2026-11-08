"""Fill and send only a verified QQ friend application form."""

import re

from task_control import InterruptibleTime, check_stop


time = InterruptibleTime()
_FORM_LABEL = '申请加好友'
_MESSAGE_LABEL = '填写验证信息'
_APPLICATION_TITLES = {'加好友/群', '加好友', '添加好友', '申请加好友'}
_DIAGNOSTIC_TITLES = _APPLICATION_TITLES | {'QQ', '综合搜索', '全网搜索'}
_DIAGNOSTIC_LABELS = (
    _FORM_LABEL, _MESSAGE_LABEL, '备注', '分组', '验证信息', '验证问题', '无需验证',
)
_DIAGNOSTIC_EDIT_CLASSES = {'q-textarea__inner', 'q-input__inner'}
_DIAGNOSTIC_BUTTONS = {'发送', '取消', '完成', '确定', '添加', '加入', '申请加入', '搜索', '关闭'}


class NoVerificationField(RuntimeError):
    """The exact target has an application form with no verification input."""

    def __init__(self, window):
        self.window = window
        super().__init__('目标号码的好友申请表单没有验证信息输入框，未填写或发送好友申请。')


def _normalized(value):
    return value.replace('\r\n', '\n')


def _validate_message(message):
    if not isinstance(message, str) or not message.strip():
        raise ValueError('好友验证信息不能为空，请在页面中填写验证信息。')


def _validate_expected_number(expected_number):
    if expected_number is not None and (
        not isinstance(expected_number, str) or re.fullmatch(r'[0-9]{5,12}', expected_number) is None
    ):
        raise ValueError('待核对的 QQ 号码格式不正确。')


def _visible_text_controls(window):
    for control in window.descendants():
        try:
            if control.element_info.control_type == 'Text' and control.is_visible():
                yield control, control.window_text().strip()
        except Exception:
            # One stale, unrelated text node must not hide the whole form.
            # Missing required labels still prevent any input or sending.
            continue


def _visible_texts(window):
    return {text for _, text in _visible_text_controls(window)}


def _form_timeout_message(desktop, process_id):
    """Describe structure only; never include names, edit values or chat text."""
    snapshots = []
    try:
        windows = desktop.windows()
    except Exception:
        return '未找到所选 QQ 的好友申请表单，未发送好友申请。诊断：无法读取窗口列表。'
    for window in windows:
        check_stop()
        try:
            if window.process_id() != process_id or not window.is_visible():
                continue
            title = window.window_text().strip()
            safe_title = title if title in _DIAGNOSTIC_TITLES else '其他窗口'
            texts = _visible_texts(window)
            classes, buttons = set(), set()
            for control in window.descendants():
                try:
                    if not control.is_visible():
                        continue
                    kind = control.element_info.control_type
                    if kind == 'Edit':
                        known = set(control.element_info.class_name.split()) & _DIAGNOSTIC_EDIT_CLASSES
                        classes.update(known or {'其他Edit类'})
                    elif kind == 'Button':
                        name = control.window_text().strip()
                        if name in _DIAGNOSTIC_BUTTONS:
                            buttons.add(name)
                except Exception:
                    continue
            candidate = title in _APPLICATION_TITLES or _FORM_LABEL in texts
            labels = '、'.join(label for label in _DIAGNOSTIC_LABELS if label in texts) or '无'
            detail = (f'{safe_title}[标签={labels}；Edit={"、".join(sorted(classes)) or "无"}；'
                      f'按钮={"、".join(sorted(buttons)) or "无"}]')
            snapshots.append((candidate, _FORM_LABEL in texts, _MESSAGE_LABEL in texts, detail))
        except Exception:
            # Diagnostics must not leak exception text from an inaccessible control.
            continue
    if any(has_form and not has_message for _, has_form, has_message, _ in snapshots):
        reason = '检测到好友申请页，但未显示“填写验证信息”标签，未发送好友申请。'
    elif any(candidate for candidate, _, _, _ in snapshots):
        reason = '检测到加好友窗口，但未识别到完整的好友申请标签，未发送好友申请。'
    else:
        reason = '未找到所选 QQ 的好友申请表单（未检测到申请页），未发送好友申请。'
    snapshots.sort(key=lambda item: not item[0])
    details = '；'.join(item[3] for item in snapshots[:6]) or '未检测到所选 QQ 的可见窗口'
    if len(snapshots) > 6:
        details += '；其余窗口已省略'
    return reason + '诊断：' + details


def _validate_form(window, process_id, *, expected_number=None):
    try:
        if window.process_id() != process_id or not window.is_visible():
            raise RuntimeError('好友申请表单已关闭或不属于所选 QQ，已停止发送。')
        texts = _visible_texts(window)
        if not {_FORM_LABEL, _MESSAGE_LABEL}.issubset(texts):
            raise RuntimeError('当前窗口不是完整的好友申请表单，已停止发送。')
        if expected_number is not None and expected_number not in texts:
            raise RuntimeError('好友申请表单目标号码不匹配，已停止填写和发送。')
    except RuntimeError:
        raise
    except Exception as error:
        raise RuntimeError('无法读取好友申请表单，已停止发送。') from error


def _has_verification_textarea(window):
    """Return only after classifying every visible Edit control safely."""
    found = False
    try:
        for control in window.descendants():
            try:
                if control.element_info.control_type != 'Edit' or not control.is_visible():
                    continue
                if 'q-textarea__inner' in control.element_info.class_name.split():
                    found = True
            except Exception as error:
                raise RuntimeError('无法确认申请表单是否有验证输入框，已停止操作。') from error
    except RuntimeError:
        raise
    except Exception as error:
        raise RuntimeError('无法确认申请表单是否有验证输入框，已停止操作。') from error
    return found


def _is_unconfigured_application(window, process_id, expected_number):
    if expected_number is None:
        return False
    try:
        if window.process_id() != process_id or not window.is_visible():
            return False
        if window.window_text().strip() != '加好友/群':
            return False
        texts = _visible_texts(window)
        return (
            _FORM_LABEL in texts
            and expected_number in texts
            and _MESSAGE_LABEL not in texts
            and not _has_verification_textarea(window)
        )
    except RuntimeError:
        raise
    except Exception:
        return False


def _verification_edit(window, process_id, *, expected_number=None):
    _validate_form(window, process_id, expected_number=expected_number)
    try:
        labels = [control for control, text in _visible_text_controls(window) if text == _MESSAGE_LABEL]
        if len(labels) != 1:
            raise RuntimeError('好友申请表单中存在多个验证信息区域，已停止发送。')
        # QQ puts verification, remark and group in separate groups. Looking
        # only at this label's direct siblings prevents writing a remark.
        edits = [
            control for control in labels[0].parent().children()
            if control.element_info.control_type == 'Edit' and control.is_visible()
        ]
        if len(edits) != 1:
            raise RuntimeError('未找到唯一的好友验证信息输入框，已停止发送。')
        edit = edits[0]
        if 'q-textarea__inner' not in edit.element_info.class_name.split():
            raise RuntimeError('好友验证信息输入框类型与预期不符，已停止发送。')
        if not edit.is_enabled():
            raise RuntimeError('好友验证信息输入框当前不可编辑，已停止发送。')
        return edit
    except RuntimeError:
        raise
    except Exception as error:
        raise RuntimeError('无法准确定位好友验证信息输入框，已停止发送。') from error


def _read_message(window, process_id, *, expected_number=None):
    edit = _verification_edit(window, process_id, expected_number=expected_number)
    try:
        value = edit.get_value()
        if not isinstance(value, str):
            raise TypeError('verification value is not text')
        return _normalized(value)
    except Exception as error:
        raise RuntimeError('无法回读好友验证信息，已停止发送。') from error


def prepare_verification_form(desktop, process_id, message, timeout=5, *, expected_number=None):
    """Wait for one application form, fill its verification and read it back.

    This function never clicks Send. Unsupported, ambiguous and truncated
    fields fail closed; no keyboard fallback can type into another field.
    """
    _validate_message(message)
    _validate_expected_number(expected_number)
    deadline = time.monotonic() + max(0, timeout)
    saw_target_mismatch = False
    while True:
        check_stop()
        forms = []
        for window in desktop.windows():
            try:
                if window.process_id() != process_id or not window.is_visible():
                    continue
                texts = _visible_texts(window)
                if (expected_number is not None
                        and window.window_text().strip() == '加好友/群'
                        and _FORM_LABEL in texts
                        and expected_number not in texts):
                    saw_target_mismatch = True
                    continue
                if _is_unconfigured_application(window, process_id, expected_number):
                    raise NoVerificationField(window)
                if {_FORM_LABEL, _MESSAGE_LABEL}.issubset(texts):
                    if expected_number is not None and expected_number not in texts:
                        saw_target_mismatch = True
                        continue
                    forms.append(window)
            except NoVerificationField:
                raise
            except Exception:
                # A closing, unrelated top-level window may be stale.
                continue
        if len(forms) > 1:
            raise RuntimeError('所选 QQ 有多个好友申请表单，请关闭多余表单后重试。')
        if forms:
            form = forms[0]
            break
        if time.monotonic() >= deadline:
            if saw_target_mismatch:
                raise RuntimeError(
                    '等待期间检测到好友申请表单目标号码不匹配，未填写或发送好友申请。'
                    + _form_timeout_message(desktop, process_id)
                )
            raise RuntimeError(_form_timeout_message(desktop, process_id))
        time.sleep(0.1)

    edit = _verification_edit(form, process_id, expected_number=expected_number)
    try:
        edit.set_edit_text(message)
    except Exception as error:
        raise RuntimeError('无法填写好友验证信息，未发送好友申请。') from error

    # The UIA setter may return before QQ updates its accessible value.
    read_deadline = time.monotonic() + min(1.0, max(0, timeout))
    while True:
        check_stop()
        if _read_message(form, process_id, expected_number=expected_number) == _normalized(message):
            return form
        if time.monotonic() >= read_deadline:
            raise RuntimeError('好友验证信息回读与配置不一致，可能被截断，未发送好友申请。')
        time.sleep(0.1)


def _is_no_pattern_error(error):
    try:
        from pywinauto.uia_defines import NoPatternInterfaceError
    except ImportError:
        return False
    return isinstance(error, NoPatternInterfaceError)


def _is_unavailable_element(error):
    try:
        from pywinauto.controls.hwndwrapper import InvalidWindowHandle
        if isinstance(error, InvalidWindowHandle):
            return True
    except ImportError:
        pass
    try:
        from comtypes import COMError
        return isinstance(error, COMError) and error.hresult & 0xffffffff == 0x80040201
    except ImportError:
        return False


def _native_window_exists(handle):
    import win32gui
    return bool(win32gui.IsWindow(handle))


def _wait_form_closed(window, original_handle, timeout, action='发送'):
    deadline = time.monotonic() + max(0, timeout)
    while True:
        check_stop()
        if original_handle:
            try:
                if not _native_window_exists(original_handle):
                    return
            except Exception:
                # Failure to inspect an HWND is not evidence that it closed.
                pass
        try:
            if not window.is_visible():
                return
        except Exception as error:
            if _is_unavailable_element(error):
                return
            raise RuntimeError(f'已触发{action}，但无法确认申请窗口是否关闭，请核对 QQ 中的申请状态。') from error
        if time.monotonic() >= deadline:
            raise RuntimeError(f'已触发{action}但申请窗口尚未关闭，请核对 QQ 中的申请状态；程序不会再次触发操作。')
        time.sleep(0.1)


def click_verified_send(window, process_id, message, send_button_text='发送', *, expected_number=None, timeout=5):
    """Invoke this verified form once, then confirm it is hidden or gone."""
    _validate_message(message)
    _validate_expected_number(expected_number)
    check_stop()
    if _read_message(window, process_id, expected_number=expected_number) != _normalized(message):
        raise RuntimeError('好友验证信息已变化，与配置不一致，已停止发送。')
    try:
        buttons = [
            control for control in window.descendants()
            if control.element_info.control_type == 'Button'
            and control.window_text().strip() == send_button_text
            and control.is_visible()
            and control.is_enabled()
        ]
        if len(buttons) != 1:
            raise RuntimeError('好友申请表单中未找到唯一可用的发送按钮，已停止发送。')
        check_stop()
        if _read_message(window, process_id, expected_number=expected_number) != _normalized(message):
            raise RuntimeError('好友验证信息已变化，与配置不一致，已停止发送。')
        original_handle = getattr(window, 'handle', None)
        button = buttons[0]
        fallback = False
        try:
            invoke = button.iface_invoke
        except AttributeError:
            fallback = True
        except Exception as error:
            if not _is_no_pattern_error(error):
                raise
            fallback = True
        if fallback:
            window.set_focus()
            check_stop()
            if _read_message(window, process_id, expected_number=expected_number) != _normalized(message):
                raise RuntimeError('好友验证信息已变化，与配置不一致，已停止发送。')
            button.click_input()
        else:
            # Do not wrap Invoke in the missing-pattern handler: an exception
            # after invocation begins can mean the request was already sent.
            invoke.Invoke()
    except RuntimeError:
        raise
    except Exception as error:
        raise RuntimeError('好友申请表单发送过程未完成，请核对 QQ 中的申请状态；程序不会重复触发发送。') from error
    _wait_form_closed(window, original_handle, timeout)
    return True


def _validate_unconfigured_form(window, process_id, expected_number):
    _validate_expected_number(expected_number)
    try:
        if window.process_id() != process_id or not window.is_visible():
            raise RuntimeError('好友申请表单已关闭或不属于所选 QQ，不能取消。')
        if window.window_text().strip() != '加好友/群':
            raise RuntimeError('当前窗口不是已确认的“加好友/群”申请页，不能取消。')
        texts = _visible_texts(window)
        if _FORM_LABEL not in texts or expected_number not in texts:
            raise RuntimeError('好友申请表单目标号码不匹配，不能取消。')
        if _MESSAGE_LABEL in texts or _has_verification_textarea(window):
            raise RuntimeError('申请表单存在验证信息区域，不能按“无验证框”方式取消。')
        return [
            control for control in window.descendants()
            if control.element_info.control_type == 'Button'
            and control.window_text().strip() == '取消'
            and control.is_visible()
            and control.is_enabled()
        ]
    except RuntimeError:
        raise
    except Exception as error:
        raise RuntimeError('无法安全核对无验证信息的申请表单，未取消。') from error


def cancel_unconfigured_form(window, process_id, expected_number, timeout=5):
    """Cancel only the same-number application form confirmed to have no textarea."""
    buttons = _validate_unconfigured_form(window, process_id, expected_number)
    if len(buttons) != 1:
        raise RuntimeError('无验证框申请页中未找到唯一可用的“取消”按钮，未执行操作。')
    check_stop()
    original_handle = getattr(window, 'handle', None)
    try:
        invoke = buttons[0].iface_invoke
    except AttributeError as error:
        raise RuntimeError('“取消”按钮不支持安全的 UIA Invoke，未执行操作。') from error
    except Exception as error:
        if not _is_no_pattern_error(error):
            raise RuntimeError('无法确认“取消”按钮的 UIA Invoke，未执行操作。') from error
        raise RuntimeError('“取消”按钮不支持安全的 UIA Invoke，未执行操作。') from error
    check_stop()
    # Never fall back to coordinates: invoking Send by mistake would submit a request.
    try:
        invoke.Invoke()
    except Exception as error:
        raise RuntimeError('无法确认无验证框申请页是否已取消，请核对 QQ；不会点击发送。') from error
    _wait_form_closed(window, original_handle, timeout, action='取消')
    return True
