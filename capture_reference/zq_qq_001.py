"""自动检查最多三个 QQ 聊天并按规则逐页点击群成员。"""  # 模块用途说明。

from __future__ import annotations  # 启用延迟类型注解求值。

import json
import os
import platform  # 读取当前操作系统信息。

import statistics
import re  # 提供群标题正则匹配功能。
import sys  # 读取 Python 运行环境信息。
import time  # 提供等待和超时计时功能。
from pathlib import Path  # 定位号码工作簿。
from task_control import atomic_save, check_pause, check_stop, set_pause_flush
from capture_progress import CaptureProgress, CaptureStorageError


WAIT_SECONDS = 1.0  # 设置界面操作之间的等待时间。
CONTINUE_WHEN_READY = True
MEMBER_CLICK_DELAY = 1.0  # 每次点击群成员后的等待时间。
PROFILE_READ_TIMEOUT = 2.0  # 等待资料卡文本稳定的最长时间。
PROFILE_POLL_INTERVAL = 0.25  # 资料卡 UIA 轮询间隔。
MAX_NUMBER_RECOGNITION_ATTEMPTS = 3  # 每位成员号码识别的最多尝试次数。
JOURNAL_FSYNC_BATCH_SIZE = 30  # 追加日志累计多少条后执行一次磁盘刷盘，可自行修改。
MEMBER_PAGE_SIZE = 14  # 每次翻页目标推进的成员行数。
MEMBER_TRIGGER_POSITION = 13  # 兼容旧配置；自动分页已不依赖此门槛。
MAX_MEMBER_PAGE_TURNS = 30  # 每个群聊最多向下翻页次数，可按需调整。
MAX_SCROLLS = 500  # 限制加入群聊列表最大滚动次数。
SESSION_SCROLL_STEP_RATIO = 0.60  # 使用 QQ 滚动接口时每次推进的可视区域比例，保留页面重叠。
SESSION_SCROLL_WHEEL_NOTCHES = 6  # 不支持滚动接口时的滚轮步数，避免一次跳过多个群聊。
MAX_MEMBERS_PER_GROUP = 2
SESSION_RETRY_LIMIT = 3  # 群列表重排导致临时定位失败时的最大重试次数，可自行修改。
GROUP_HEADER_PATTERN = re.compile(r"\(\s*\d+\s*\)\s*$")  # 匹配群标题末尾的成员数量。
MEMBER_COUNT_PATTERN = re.compile(r"群聊成员\s*(\d+)")  # 读取成员面板显示的总人数。
GROUP_MEMBER_LABEL_PATTERN = re.compile(r"^\(?\s*\d+\s*人\s*\)?$")
GROUP_NUMBER_VALUE_PATTERN = re.compile(r"^\d{5,12}$")
EXCLUDED_CHAT_NAMES = {
    "AI资源共享群",
    "tokkenShop中转站",
}  # 精确匹配并跳过的群聊名称，不计入检查名额。
EXCLUDED_MEMBER_NAMES = {"Q群管家"}  # 按成员显示名称精确匹配的跳过名单。
NUMBER_WORKBOOK_PATH = Path(__file__).resolve().with_name("qq_member_numbers.xlsx")
NUMBER_JOURNAL_PATH = Path(__file__).resolve().with_name("qq_member_numbers.jsonl")
NUMBER_TEXT_PATTERN = re.compile(r"^\s*(\d{5,12})\s*$")
WINDOW_TITLE = "QQ"
SESSION_LIST_TEXT = "会话列表"
MEMBER_LIST_TEXT = "成员列表"
TARGET_QQ_NUMBER = "3805745874"  # 填写要操作的 QQ 号；留空时脚本会提示并停止。
SELECTED_PROCESS_ID = None
AVATAR_BUTTON_TEXT = "昵称"
AVATAR_AUTO_DETECT = True
AVATAR_TIMEOUT_SECONDS = 5.0
USE_OCR = True

_journal_pending_sync_count = 0
_capture_workbook = None
_capture_worksheet = None
_capture_progress = None


def flush_capture():
    """Flush the journal before saving Excel, including on pause or stop."""
    try:
        sync_journal()
        if _capture_workbook is not None:
            reconcile_journal(_capture_workbook, _capture_worksheet, load_journal_records(), force_save=True)
    except Exception as error:
        raise CaptureStorageError(f'无法保存采集结果，请关闭 Excel/WPS 后重试：{error}') from error


def group_sheet_title(group_name, group_number, suffix=''):
    name = re.sub(r'[\\/*?:\[\]\x00-\x1f]', '_', str(group_name)).strip(" '") or 'QQ群'
    tail = '_' + str(group_number) + suffix
    # Excel limits sheet names to 31 UTF-16 code units; keep the full group ID.
    budget = 31 - len(tail)
    while len(name.encode('utf-16-le')) // 2 > budget:
        name = name[:-1]
    return name + tail


def sheet_group_headers(sheet):
    return [(row[0].row, str(row[0].value).strip(), str(row[1].value).strip())
            for row in sheet.iter_rows(min_col=1, max_col=2)
            if row[0].value is not None and row[1].value is not None
            and GROUP_NUMBER_VALUE_PATTERN.fullmatch(str(row[1].value).strip())]


def sheet_member_numbers(sheet):
    return {str(row[0]).strip() for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True)
            if row[0] is not None and NUMBER_TEXT_PATTERN.fullmatch(str(row[0]).strip())}


def ensure_group_sheet(workbook, summary, group_name, group_number):
    """Match by group ID, including an existing first sheet; names may change."""
    group_number = str(group_number)
    for sheet in workbook.worksheets:
        headers = sheet_group_headers(sheet)
        if len(headers) == 1 and headers[0][0] == 1 and headers[0][2] == group_number:
            return sheet, False
    title = group_sheet_title(group_name, group_number)
    counter = 2
    while title.casefold() in {name.casefold() for name in workbook.sheetnames}:
        title = group_sheet_title(group_name, group_number, f'_{counter}')
        counter += 1
    empty = (len(workbook.worksheets) == 1 and summary.max_row == 1
             and all(cell.value is None for row in summary for cell in row))
    if empty:
        sheet = summary
        sheet.title = title
    else:
        sheet = workbook.create_sheet(title)
    sheet.cell(1, 1, str(group_name)).data_type = 's'
    sheet.cell(1, 2, group_number).number_format = '@'
    sheet.column_dimensions['A'].width = 28
    sheet.column_dimensions['B'].width = 20
    return sheet, True


def migrate_legacy_group_blocks(workbook):
    """Copy old mixed-group blocks into group sheets; keep the source as history."""
    from copy import copy
    changed = False
    for source in list(workbook.worksheets):
        headers = sheet_group_headers(source)
        if len(headers) < 2:
            continue
        for index, (start, name, number) in enumerate(headers):
            target, created = ensure_group_sheet(workbook, source, name, number)
            changed = created or changed
            existing = {str(row[0].value): row[0].row for row in target.iter_rows(min_row=2)
                        if row[0].value is not None}
            end = headers[index + 1][0] if index + 1 < len(headers) else source.max_row + 1
            for row in source.iter_rows(min_row=start + 1, max_row=end - 1):
                key = str(row[0].value).strip() if row[0].value is not None else ''
                if not NUMBER_TEXT_PATTERN.fullmatch(key):
                    continue
                destination = existing.get(key)
                if destination is None:
                    destination = target.max_row + 1
                    existing[key] = destination
                for cell in row:
                    if cell.value is None:
                        continue
                    out = target.cell(destination, cell.column)
                    if out.value is not None:
                        continue
                    out.value = cell.value
                    out._style = copy(cell._style)
                    if cell.comment:
                        out.comment = copy(cell.comment)
                    if cell.hyperlink:
                        out._hyperlink = copy(cell.hyperlink)
                    changed = True
        if not source.title.startswith('历史汇总'):
            title = '历史汇总'
            index = 2
            while title in workbook.sheetnames:
                title = f'历史汇总_{index}'
                index += 1
            source.title = title
            changed = True
    return changed


def configure_number_storage(workbook_path):
    """Each selected workbook owns its journal; never import a neighboring file."""
    global NUMBER_JOURNAL_PATH
    NUMBER_JOURNAL_PATH = Path(workbook_path).with_suffix('.jsonl')


def resume_for_new_task(position):
    # Completed groups are filtered before opening their member list.
    if not position:
        return {'page_names': ['__new_task_from_top__'], 'next_index': 0,
                'page_turns': 0, 'scroll_percent': 0.0, 'saved_count': 0}
    return position


def completed_groups_for_resume(journal_records, progress):
    """Use explicit completion only; the account checkpoint wins over old logs."""
    completed = set()
    accounts = progress.data.get('accounts', {})
    legacy_account_is_unambiguous = not (set(accounts) - {progress.account})
    for record in journal_records:
        account = record.get('account')
        if account is not None and str(account) != progress.account:
            continue
        if account is None and not legacy_account_is_unambiguous:
            continue
        group_number = str(record['group_number'])
        if record['type'] == 'group_complete':
            completed.add(group_number)
        elif record['type'] in {'group', 'member'}:
            completed.discard(group_number)
    for group_number, position in progress.groups.items():
        if position.get('completed'):
            completed.add(str(group_number))
        else:
            completed.discard(str(group_number))
    return completed


def append_to_group_sheet(sheet, number):
    if any(str(row[0]) == number for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True)):
        return False
    sheet.cell(sheet.max_row + 1, 1, number).number_format = '@'
    return True


def capture_page_position(main_window, work_area):
    panel = get_visible_member_panel(main_window)
    members = get_member_list_window(main_window)
    if panel is None or members is None:
        return [], None
    names = [member_item_text(item) for item in collect_member_items(members, panel, work_area)]
    try:
        percent = float(panel.iface_scroll.CurrentVerticalScrollPercent)
        if not 0 <= percent <= 100:
            percent = None
    except Exception:
        percent = None
    return names, percent


def restore_capture_position(main_window, desktop, pyautogui, work_area, resume):
    """Verify the entire visible page before trusting a persisted row index."""
    if not resume.get('page_names'):
        return 0, 0
    panel = get_visible_member_panel(main_window)
    percent = resume.get('scroll_percent')
    positioned = False
    if percent is not None and panel is not None:
        try:
            panel.iface_scroll.SetScrollPercent(-1.0, percent)
            positioned = bool(wait_for_resume_state(
                lambda: capture_page_position(main_window, work_area)[0] == resume['page_names'],
                WAIT_SECONDS, interval=0.15))
        except Exception:
            pass
    names, _ = capture_page_position(main_window, work_area)
    if names == resume['page_names'] and (positioned or resume.get('page_turns', 0) == 0):
        print(f"已恢复本群第 {resume.get('page_turns', 0) + 1} 页，下一个成员行 {resume['next_index'] + 1}。")
        return resume['next_index'], resume.get('page_turns', 0)
    # A changed QQ layout/list must never cause unchecked rows to be skipped.
    if panel is not None:
        old_names, old_percent = capture_page_position(main_window, work_area)
        try:
            panel.iface_scroll.SetScrollPercent(-1.0, 0.0)
        except Exception:
            rect = panel.rectangle()
            x = min(work_area[2] - 1, max(work_area[0], (rect.left + rect.right) // 2))
            y = min(work_area[3] - 1, max(work_area[1], (rect.top + rect.bottom) // 2))
            pyautogui.moveTo(x, y)
            pyautogui.scroll(1000, x=x, y=y)
        last_names = None
        def reset_ready():
            nonlocal last_names
            names, current_percent = capture_page_position(main_window, work_area)
            if current_percent is not None:
                return (bool(names) and current_percent <= 0.01
                        and (old_percent == 0 or names != old_names))
            settled = bool(names) and last_names == names
            last_names = names
            return settled
        if not wait_for_resume_state(reset_ready, WAIT_SECONDS, interval=0.15):
            raise RuntimeError('成员列表返回顶部后尚未就绪，保留进度，请重试。')
    if percent is None:
        for page in range(resume.get('page_turns', 0) + 1):
            check_pause()
            main_window = find_main_qq_window(desktop)
            names, _ = capture_page_position(main_window, work_area)
            if names == resume['page_names']:
                print('已重新定位保存的成员页，继续上次未完成的位置。')
                return resume['next_index'], page
            if page < resume.get('page_turns', 0):
                advanced, _ = scroll_member_list(desktop, pyautogui, work_area)
                if not advanced:
                    break
        # Reset once more after unsuccessful searching.
        return restore_capture_position(main_window, desktop, pyautogui, work_area,
                                        dict(resume, scroll_percent=0, page_names=['__unmatched__']))
    print('成员列表已变化，重新核对成员并按 QQ 号码去重，已保存数据保留。')
    return 0, 0


def save_member_position(group_section, main_window, work_area, page_names, next_index, page_turns, saved_count):
    progress = group_section.get('progress')
    if progress is None:
        return
    panel = get_visible_member_panel(main_window)
    try:
        percent = float(panel.iface_scroll.CurrentVerticalScrollPercent)
        if not 0 <= percent <= 100:
            percent = None
    except Exception:
        percent = None
    try:
        sync_journal()
        progress.save(group_section['group_number'], group_name=group_section['group_name'],
                      page_names=page_names, next_index=next_index, page_turns=page_turns,
                      scroll_percent=percent, saved_count=saved_count, completed=False)
    except Exception as error:
        raise CaptureStorageError(f'无法保存成员采集进度，已停止：{error}') from error


def group_saved_count(worksheet, group_number):
    section = find_group_section_in_sheet(worksheet, group_number)
    if section is None:
        return 0
    count = 0
    for row in range(section['group_row'] + 1, worksheet.max_row + 1):
        value = worksheet.cell(row, 1).value
        group = worksheet.cell(row, 2).value
        if value is not None and GROUP_NUMBER_VALUE_PATTERN.fullmatch(str(group).strip()):
            break
        if NUMBER_TEXT_PATTERN.fullmatch(str(value).strip()) and not group:
            count += 1
    return count


def wait_for_resume_state(probe, timeout, interval=0.15):
    """Confirm only the saved list position; retain normal capture wait settings."""
    deadline = time.monotonic() + timeout
    while True:
        check_stop()
        result = probe()
        if result:
            return result
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return None
        time.sleep(min(interval, remaining))


def wait_for_result(probe, timeout, interval=None):
    """Check immediately, then poll only while incomplete; never extend the deadline."""
    deadline = time.monotonic() + timeout
    interval = PROFILE_POLL_INTERVAL if interval is None else interval
    while True:
        check_stop()
        result = probe()
        if result:
            return result
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return None
        time.sleep(min(interval, remaining))


def after_ui_action(probe, previous_wait, *, timeout=None, interval=None):
    """Retain the saved fixed-delay mode and the original wait/timeout budgets."""
    if not CONTINUE_WHEN_READY:
        time.sleep(previous_wait)
        return probe()
    return wait_for_result(probe, previous_wait if timeout is None else timeout, interval)


def reset_joined_groups_to_top(desktop, main_window, pyautogui, work_area):
    """Confirm both scroll position and visible rows before scanning the first page."""
    pane = get_contact_group_list(main_window)
    if pane is None:
        raise RuntimeError('无法定位群聊列表，不能确认列表顶部。')
    before = collect_joined_group_rows(main_window, work_area)
    def signature(rows):
        return tuple((name, row.rectangle().top, row.rectangle().bottom) for name, row in rows)
    baseline = signature(before)
    def top_marker_visible(current_pane):
        pane_rect = current_pane.rectangle()
        for control in current_pane.descendants(control_type='Button'):
            if control.window_text().strip().startswith('置顶群聊'):
                rect = control.rectangle()
                return (control.is_visible() and rect.height() > 0
                        and abs(rect.top - pane_rect.top) <= 3
                        and point_in_work_area((rect.left + rect.right) // 2,
                                               (rect.top + rect.bottom) // 2, work_area))
        return False
    if before and top_marker_visible(pane):
        return main_window
    already_top = False
    try:
        scroll = pane.iface_scroll
        if not scroll.CurrentVerticallyScrollable:
            return main_window
        already_top = float(scroll.CurrentVerticalScrollPercent) <= 0.01
        if already_top and before:
            return main_window
        scroll.SetScrollPercent(-1.0, 0.0)
    except Exception:
        if not before:
            raise RuntimeError('无法确认群聊列表位于顶部：没有可操作的群行。')
        rect = before[0][1].rectangle()
        x, y = (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2
        pyautogui.moveTo(x, y)
        pyautogui.scroll(100, x=x, y=y)
    last_signature = None
    def ready():
        nonlocal last_signature
        current = find_main_qq_window(desktop)
        current_pane = get_contact_group_list(current) if current else None
        rows = collect_joined_group_rows(current, work_area) if current_pane else []
        if not rows:
            return None
        current_signature = signature(rows)
        try:
            scroll = current_pane.iface_scroll
            at_top = (not scroll.CurrentVerticallyScrollable
                      or float(scroll.CurrentVerticalScrollPercent) <= 0.01)
            return current if at_top and (already_top or current_signature != baseline) else None
        except Exception:
            if top_marker_visible(current_pane):
                return current
            # Without a scroll interface, require two matching nonempty observations.
            settled = last_signature == current_signature
            last_signature = current_signature
            return current if settled else None
    if not CONTINUE_WHEN_READY:
        time.sleep(min(0.5, WAIT_SECONDS))
        return find_main_qq_window(desktop)
    result = after_ui_action(ready, min(0.5, WAIT_SECONDS), interval=0.05)
    if result is None:
        raise RuntimeError('群聊列表返回顶部后尚未就绪，请重试。')
    return result


def wait_for_member_panel(desktop, work_area):
    refreshed = False
    def ready():
        nonlocal refreshed
        current = find_main_qq_window(desktop)
        if current is None:
            return None
        members = get_member_list_window(current)
        panel = get_visible_member_panel(current)
        if members is not None and panel is not None:
            if collect_member_items(members, panel, work_area):
                return current
            # QQ may reset scroll position while retaining off-screen virtual
            # rows from the previous group. Refresh without changing position.
            if not refreshed:
                refreshed = True
                try:
                    scroll = panel.iface_scroll
                    percent = float(scroll.CurrentVerticalScrollPercent)
                    if scroll.CurrentVerticallyScrollable and 0 <= percent <= 100:
                        target = percent + .01 if percent < 99.99 else percent - .01
                        try:
                            scroll.SetScrollPercent(-1.0, target)
                        finally:
                            scroll.SetScrollPercent(-1.0, percent)
                        print('成员区域尚未刷新，已保持原位置刷新列表。')
                        if collect_member_items(members, panel, work_area):
                            return current
                except Exception:
                    pass
        return None
    return after_ui_action(ready, WAIT_SECONDS, interval=0.15)


def load_dependencies():  # 加载桌面自动化所需依赖。
    """加载 Windows UIA 和屏幕滚动依赖；缺少时给出安装提示。"""
    try:  # 尝试导入已安装的桌面自动化库。
        import pyautogui  # 导入鼠标滚动控制库。
        import win32api  # 导入显示器和工作区查询接口。
        import win32con  # 导入 Windows 窗口常量。
        import win32gui  # 导入 Windows 窗口控制接口。
        from pywinauto import Desktop  # 导入 UI Automation 桌面入口。
        from openpyxl import Workbook, load_workbook  # 导入 XLSX 读写接口。
    except ImportError as error:  # 捕获缺少依赖的情况。
        print(f"缺少依赖：{error.name or error}")  # 输出缺少的包名。
        print("请在命令行运行：python -m pip install pywinauto pywin32 pyautogui openpyxl")  # 提供必需依赖安装命令。
        raise SystemExit(2) from error  # 以错误状态结束程序。
    try:
        if not USE_OCR:
            raise ImportError("OCR disabled")
        from rapidocr_onnxruntime import RapidOCR
    except Exception:
        RapidOCR = None
    return (
        Desktop,
        pyautogui,
        win32api,
        win32con,
        win32gui,
        Workbook,
        load_workbook,
        RapidOCR,
    )


# 根据 QQ 所在显示器取得工作区边界，供鼠标操作避开任务栏。
def get_work_area(win32api, win32con, window_handle):
    """获取 QQ 所在显示器的工作区；系统任务栏不在此区域内。"""
    monitor = win32api.MonitorFromWindow(
        window_handle, win32con.MONITOR_DEFAULTTONEAREST
    )
    return win32api.GetMonitorInfo(monitor)["Work"]


# 判断给定屏幕坐标是否落在 Windows 工作区内。
def point_in_work_area(x, y, work_area):
    left, top, right, bottom = work_area
    return left <= x < right and top <= y < bottom


# 将控件可见区域裁到工作区后再点击，完全越界时不执行鼠标操作。
def click_control_in_work_area(control, work_area, clip_rect=None, *, avatar=False):
    rect = control.rectangle()
    if rect.width() <= 0 or rect.height() <= 0:
        return False
    work_left, work_top, work_right, work_bottom = work_area
    left = max(rect.left, work_left)
    top = max(rect.top, work_top)
    right = min(rect.right, work_right)
    bottom = min(rect.bottom, work_bottom)
    if clip_rect is not None:
        left = max(left, clip_rect.left)
        top = max(top, clip_rect.top)
        right = min(right, clip_rect.right)
        bottom = min(bottom, clip_rect.bottom)
    if right <= left or bottom <= top:
        return False
    x = min(right - 1, max(left, rect.left + 20)) if avatar else (left + right) // 2
    y = (top + bottom) // 2
    control.click_input(coords=(x - rect.left, y - rect.top))
    return True


# 识别属于 QQ 进程或由 QQ 主窗口拥有的弹出窗口。
def is_qq_related_window(window, process_id, main_handle, win32gui, win32con):
    try:
        if window.process_id() == process_id:
            return True
        owner = win32gui.GetWindow(window.handle, win32con.GW_OWNER)
        visited = set()
        while owner and owner not in visited:
            if owner == main_handle:
                return True
            visited.add(owner)
            owner = win32gui.GetWindow(owner, win32con.GW_OWNER)
    except Exception:
        return False
    return False


# 收集 QQ 及其弹窗的 UIA 文本、控件类型和矩形，用于前后状态比对。
def member_profile_windows(desktop, process_id, main_handle, win32gui):
    """Use native window metadata before asking UIA for a specific profile."""
    import win32process
    handles = []
    def visit(handle, unused):
        if (handle != main_handle and win32gui.IsWindowVisible(handle)
                and win32process.GetWindowThreadProcessId(handle)[1] == process_id
                and win32gui.GetWindowText(handle).strip() in {'资料卡', '个人资料', '用户资料'}):
            handles.append(handle)
        return True
    win32gui.EnumWindows(visit, None)
    return [desktop.window(handle=handle).wrapper_object() for handle in handles]


def collect_qq_snapshot(desktop, process_id, main_handle, win32gui, win32con):
    """Read QQ-labelled header text first and stop as soon as its number is ready."""
    snapshot = {"controls": {}, "windows": {}}
    roots = []
    for window in member_profile_windows(desktop, process_id, main_handle, win32gui):
        roots.append((window.handle, window))
    if not roots:
        main_window = desktop.window(handle=main_handle).wrapper_object()
        for name in ('资料卡', '个人资料', '用户资料'):
            for card in main_window.descendants(title=name):
                if (card.element_info.control_type in {"Pane", "Window", "Group", "Document"}
                        and card.is_visible()):
                    roots.append((main_handle, card))
    single_card = len(roots) == 1
    empty = {"controls": {}, "windows": {}}
    for handle, root in roots:
        try:
            rect = root.rectangle()
            bounds = (rect.left, rect.top, rect.right, rect.bottom)
            snapshot["windows"][handle] = {
                "title": "资料卡", "bounds": bounds,
                "process_id": process_id, "profile": True,
            }
            # QQ-space embeds, images, buttons and Documents can make COM calls
            # stall after the number has already appeared. Do not request them.
            for kind in ("Text", "Edit"):
                for control in root.descendants(control_type=kind):
                    check_stop()
                    text = control.window_text().strip()
                    if not text or not control.is_visible():
                        continue
                    rect = control.rectangle()
                    if rect.width() <= 0 or rect.height() <= 0:
                        continue
                    bounds = (rect.left, rect.top, rect.right, rect.bottom)
                    signature = (handle, kind, text, bounds)
                    snapshot["controls"][signature] = {
                        "window_handle": handle, "window_title": "资料卡",
                        "control_type": kind, "text": text, "bounds": bounds,
                    }
                    if single_card and select_profile_qq_number(snapshot, empty) is not None:
                        return snapshot
        except Exception:
            continue
    return snapshot


def report_slow_member_stage(stage, started):
    elapsed = time.monotonic() - started
    if elapsed >= 1.0:
        print(f'采集耗时：{stage} {elapsed:.2f} 秒。', flush=True)


def select_profile_qq_number(snapshot, before_snapshot):
    """Require an explicit QQ label, and reject the preceding card's old number."""
    controls = list(snapshot["controls"].values())
    old_numbers = {number_from_text(c["text"])
                   for c in before_snapshot["controls"].values()}
    candidates = set()
    label_pattern = re.compile(r"QQ\s*(?:号|号码)?\s*[:：]?", re.IGNORECASE)
    combined_pattern = re.compile(r"QQ\s*(?:号|号码)?\s*[:：]?\s*([0-9]{5,12})", re.IGNORECASE)
    for control in before_snapshot['controls'].values():
        previous = combined_pattern.fullmatch(control['text'])
        if previous:
            old_numbers.add(previous.group(1))
    for label in controls:
        match = combined_pattern.fullmatch(label["text"])
        if match:
            candidates.add(match.group(1))
            continue
        if not label_pattern.fullmatch(label["text"]):
            continue
        left, top, right, bottom = label["bounds"]
        for value in controls:
            if value["window_handle"] != label["window_handle"]:
                continue
            number = number_from_text(value["text"])
            if number is None:
                continue
            vl, vt, vr, vb = value["bounds"]
            height = min(bottom - top, vb - vt)
            overlap = min(bottom, vb) - max(top, vt)
            if height > 0 and overlap >= height * .5 and -3 <= vl - right <= max(80, height * 6):
                candidates.add(number)
    # Geometry changes alone cannot make the previous member a new profile.
    candidates.difference_update(old_numbers)
    return next(iter(candidates)) if len(candidates) == 1 else None


# 从单个 UIA 文本控件中提取符合长度要求的号码。
def number_from_text(text):
    match = NUMBER_TEXT_PATTERN.fullmatch(text)
    if match:
        return match.group(1)
    if any(label in text for label in ("QQ号", "QQ号码", "账号", "帐号")):
        digits = re.findall(r"\d{5,12}", text)
        if len(digits) == 1:
            return digits[0]
    return None


# 规范 OCR 插入的空格或分隔符后，提取可能的号码片段。
def numbers_from_ocr_text(text):
    """Extract 5-12 digit identifiers even when OCR inserts spaces or separators."""
    compact = re.sub(r"(?<=\d)[\s,，.。:_-]+(?=\d)", "", text)
    return re.findall(r"(?<!\d)\d{5,12}(?!\d)", compact)


# 只保留点击前后发生新增或文本、位置变化的控件，过滤旧资料卡残留。
def new_snapshot_controls(before_snapshot, after_snapshot):
    return [
        item
        for signature, item in after_snapshot["controls"].items()
        if signature not in before_snapshot["controls"]
    ]


# 根据号码控件与附近昵称文本的空间关系，选择最可信的号码候选。
def select_number_below_label(controls, candidate_controls=None):
    candidate_controls = controls if candidate_controls is None else candidate_controls
    candidates = []
    for control in candidate_controls:
        numbers = numbers_from_ocr_text(control["text"])
        if not numbers:
            number = number_from_text(control["text"])
            numbers = [number] if number is not None else []
        candidates.extend((number, control) for number in numbers)
    if not candidates:
        return None, 0

    labels = [
        control
        for control in controls
        if control["text"] and number_from_text(control["text"]) is None
    ]
    ranked = []
    for number, candidate in candidates:
        left, top, right, bottom = candidate["bounds"]
        center_x = (left + right) // 2
        for label in labels:
            if label["window_handle"] != candidate["window_handle"]:
                continue
            label_left, label_top, label_right, label_bottom = label["bounds"]
            vertical_gap = top - label_bottom
            label_center_x = (label_left + label_right) // 2
            horizontal_gap = left - label_right
            vertical_distance = abs((top + bottom) // 2 - (label_top + label_bottom) // 2)
            below = -12 <= vertical_gap <= 180 and abs(center_x - label_center_x) <= 300
            beside = -24 <= horizontal_gap <= 300 and vertical_distance <= 70
            if below or beside:
                distance = (
                    max(0, vertical_gap) + abs(center_x - label_center_x)
                    if below
                    else max(0, horizontal_gap) + vertical_distance
                )
                ranked.append((distance, number))
    if not ranked:
        return None, len({number for number, _ in candidates})

    best_score_by_number = {}
    for score, number in ranked:
        best_score_by_number[number] = min(
            score, best_score_by_number.get(number, score)
        )
    ranked_numbers = sorted(
        (score, number) for number, score in best_score_by_number.items()
    )
    if len(ranked_numbers) == 1 or ranked_numbers[0][0] < ranked_numbers[1][0]:
        return ranked_numbers[0][1], len({number for number, _ in candidates})
    return None, len({number for number, _ in candidates})


# 汇总 UIA 快照差异的控件类型和数字候选数量，不输出成员文本。
def snapshot_diagnostic(before_snapshot, after_snapshot):
    controls = new_snapshot_controls(before_snapshot, after_snapshot)
    type_counts = {}
    for control in controls:
        name = control["control_type"]
        type_counts[name] = type_counts.get(name, 0) + 1
    candidates = {
        number_from_text(control["text"])
        for control in controls
        if number_from_text(control["text"]) is not None
    }
    added_windows = set(after_snapshot["windows"]) - set(before_snapshot["windows"])
    return (
        f"新增 UIA 控件={len(controls)}，类型={type_counts}，"
        f"数字候选={len(candidates)}，新增窗口={len(added_windows)}"
    )


# 在点击成员后轮询 UIA，只用发生变化的号码控件与附近文本进行匹配。
def extract_member_number_from_popup(
    desktop,
    before_snapshot,
    process_id,
    main_handle,
    clicked_rect,
    win32gui,
    win32con,
):
    """Poll QQ UIA controls briefly and choose a number positioned below a nearby label."""
    deadline = time.monotonic() + PROFILE_READ_TIMEOUT
    latest_snapshot = before_snapshot
    latest_diagnostic = "未发现点击后新增的 QQ 控件"
    while time.monotonic() <= deadline:
        check_stop()
        latest_snapshot = collect_qq_snapshot(
            desktop, process_id, main_handle, win32gui, win32con
        )
        number = select_profile_qq_number(latest_snapshot, before_snapshot)
        if number is not None:
            return number, "资料卡 QQ 标签已确认", latest_snapshot
        latest_diagnostic = snapshot_diagnostic(before_snapshot, latest_snapshot)
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        time.sleep(min(PROFILE_POLL_INTERVAL, remaining))
    return None, latest_diagnostic, latest_snapshot


# 将截图矩形裁剪到工作区内，避免请求屏幕范围之外的区域。
def clamp_region(bounds, work_area):
    left, top, right, bottom = bounds
    work_left, work_top, work_right, work_bottom = work_area
    region = (
        max(left, work_left),
        max(top, work_top),
        min(right, work_right),
        min(bottom, work_bottom),
    )
    if region[2] <= region[0] or region[3] <= region[1]:
        return None
    return region


# 结合新弹窗、点击附近控件和成员点击位置，生成候选 OCR 截图区域。
def popup_ocr_regions(before_snapshot, after_snapshot, clicked_rect, work_area, main_handle):
    regions = []
    for handle in set(after_snapshot["windows"]) - set(before_snapshot["windows"]):
        if handle == main_handle:
            continue
        bounds = after_snapshot["windows"][handle]["bounds"]
        width = bounds[2] - bounds[0]
        height = bounds[3] - bounds[1]
        if 100 <= width <= 850 and 60 <= height <= 700:
            region = clamp_region(bounds, work_area)
            if region:
                regions.append(region)

    click_x = (clicked_rect[0] + clicked_rect[2]) // 2
    click_y = (clicked_rect[1] + clicked_rect[3]) // 2
    nearby_controls = [
        item
        for item in after_snapshot["controls"].values()
        if item["bounds"][2] >= click_x - 520
        and item["bounds"][0] <= click_x + 520
        and item["bounds"][3] >= click_y - 420
        and item["bounds"][1] <= click_y + 420
    ]
    if nearby_controls:
        left = min(item["bounds"][0] for item in nearby_controls)
        top = min(item["bounds"][1] for item in nearby_controls)
        right = max(item["bounds"][2] for item in nearby_controls)
        bottom = max(item["bounds"][3] for item in nearby_controls)
        if 100 <= right - left <= 850 and 60 <= bottom - top <= 700:
            region = clamp_region((left - 24, top - 24, right + 24, bottom + 24), work_area)
            if region:
                regions.append(region)

    center_x = (clicked_rect[0] + clicked_rect[2]) // 2
    center_y = (clicked_rect[1] + clicked_rect[3]) // 2
    fallback = clamp_region(
        (center_x - 360, center_y - 220, center_x + 420, center_y + 300),
        work_area,
    )
    if fallback:
        regions.append(fallback)
    return list(dict.fromkeys(regions))


# 合并同一行被 OCR 拆开的数字片段，再按邻近昵称的位置筛选号码。
def extract_number_from_ocr(ocr_texts):
    controls = []
    numeric_fragments = []
    for text, bounds in ocr_texts:
        compact = re.sub(r"[\s,，.。:_-]+", "", text)
        if compact.isdigit() and 1 <= len(compact) <= 12:
            numeric_fragments.append((compact, bounds))
        else:
            controls.append(
                {
                    "text": text,
                    "bounds": bounds,
                    "window_handle": 0,
                    "control_type": "OCR",
                }
            )
    numeric_fragments.sort(key=lambda item: ((item[1][1] + item[1][3]) // 2, item[1][0]))
    line_runs = []
    for digits, bounds in numeric_fragments:
        left, top, right, bottom = bounds
        center_y = (top + bottom) // 2
        if (
            line_runs
            and abs(line_runs[-1]["center_y"] - center_y) <= 18
            and left - line_runs[-1]["right"] <= 36
        ):
            run = line_runs[-1]
            run["digits"] += digits
            run["right"] = right
            run["top"] = min(run["top"], top)
            run["bottom"] = max(run["bottom"], bottom)
        else:
            line_runs.append(
                {
                    "digits": digits,
                    "left": left,
                    "top": top,
                    "right": right,
                    "bottom": bottom,
                    "center_y": center_y,
                }
            )
    for run in line_runs:
        if 5 <= len(run["digits"]) <= 12:
            controls.append(
                {
                    "text": run["digits"],
                    "bounds": (run["left"], run["top"], run["right"], run["bottom"]),
                    "window_handle": 0,
                    "control_type": "OCR",
                }
            )
    return select_number_below_label(controls)


# 截取候选资料卡区域并调用 RapidOCR，作为 UIA 读取失败时的后备方案。
def read_number_with_ocr(
    pyautogui,
    ocr_state,
    before_snapshot,
    after_snapshot,
    clicked_rect,
    work_area,
    main_handle,
):
    if ocr_state["factory"] is None:
        return None, "RapidOCR 未安装"
    try:
        if ocr_state["engine"] is None:
            ocr_state["engine"] = ocr_state["factory"]()
        import cv2
        import numpy

        for region in popup_ocr_regions(
            before_snapshot, after_snapshot, clicked_rect, work_area, main_handle
        ):
            region_left, region_top, region_right, region_bottom = region
            # 将左、上、右、下边界换算成截图接口需要的左、上、宽、高。
            screenshot_region = (
                region_left,
                region_top,
                region_right - region_left,
                region_bottom - region_top,
            )
            image = pyautogui.screenshot(region=screenshot_region).convert("RGB")
            pixels = numpy.asarray(image)
            detections, _ = ocr_state["engine"](
                cv2.cvtColor(pixels, cv2.COLOR_RGB2BGR)
            )
            ocr_texts = []
            for detection in detections or []:
                box, text = detection[0], str(detection[1]).strip()
                xs = [point[0] for point in box]
                ys = [point[1] for point in box]
                bounds = (
                    region[0] + int(min(xs)),
                    region[1] + int(min(ys)),
                    region[0] + int(max(xs)),
                    region[1] + int(max(ys)),
                )
                ocr_texts.append((text, bounds))
            number, candidate_count = extract_number_from_ocr(ocr_texts)
            if number is not None:
                return number, f"OCR 确认；候选数字={candidate_count}"
        return None, "OCR 未找到与上方名称位置相邻的号码"
    except Exception as error:
        return None, f"OCR 失败：{type(error).__name__}"


def append_journal_record(record):
    """追加一条记录；按配置批量执行磁盘刷盘。"""
    global _journal_pending_sync_count

    NUMBER_JOURNAL_PATH.parent.mkdir(parents=True, exist_ok=True)
    line = json.dumps(record, ensure_ascii=False, separators=(",", ":"))
    with NUMBER_JOURNAL_PATH.open("a", encoding="utf-8", newline="\n") as stream:
        stream.write(line + "\n")
        stream.flush()
        _journal_pending_sync_count += 1
        if _journal_pending_sync_count >= max(1, int(JOURNAL_FSYNC_BATCH_SIZE)):
            os.fsync(stream.fileno())
            _journal_pending_sync_count = 0


def sync_journal():
    """强制刷出尚未达到批量阈值的追加日志记录。"""
    global _journal_pending_sync_count
    if _journal_pending_sync_count <= 0 or not NUMBER_JOURNAL_PATH.exists():
        return
    with NUMBER_JOURNAL_PATH.open("a", encoding="utf-8", newline="\n") as stream:
        stream.flush()
        os.fsync(stream.fileno())
    _journal_pending_sync_count = 0


def load_journal_records():
    """读取追加日志；发现损坏记录时停止，避免静默猜测数据。"""
    if not NUMBER_JOURNAL_PATH.exists():
        return []
    records = []
    with NUMBER_JOURNAL_PATH.open("r", encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, start=1):
            text = line.strip()
            if not text:
                continue
            try:
                record = json.loads(text)
            except json.JSONDecodeError as error:
                raise RuntimeError(
                    f"号码日志第 {line_number} 行损坏，请检查 {NUMBER_JOURNAL_PATH}。"
                ) from error
            if (not isinstance(record, dict)
                    or record.get("type") not in {"group", "member", "group_complete"}):
                raise RuntimeError(
                    f"号码日志第 {line_number} 行格式无效，请检查 {NUMBER_JOURNAL_PATH}。"
                )
            if record["type"] == "group_complete":
                if not GROUP_NUMBER_VALUE_PATTERN.fullmatch(
                        str(record.get("group_number", "")).strip()
                ):
                    raise RuntimeError(
                        f"号码日志第 {line_number} 行群完成标记无效，请检查 {NUMBER_JOURNAL_PATH}。"
                    )
            elif record["type"] == "group":
                if (not str(record.get("group_name", "")).strip()
                        or not GROUP_NUMBER_VALUE_PATTERN.fullmatch(
                            str(record.get("group_number", "")).strip()
                        )):
                    raise RuntimeError(
                        f"号码日志第 {line_number} 行群信息无效，请检查 {NUMBER_JOURNAL_PATH}。"
                    )
            elif (not GROUP_NUMBER_VALUE_PATTERN.fullmatch(
                    str(record.get("group_number", "")).strip()
                  )
                  or not NUMBER_TEXT_PATTERN.fullmatch(
                      str(record.get("number", "")).strip()
                  )):
                raise RuntimeError(
                    f"号码日志第 {line_number} 行成员信息无效，请检查 {NUMBER_JOURNAL_PATH}。"
                )
            records.append(record)
    return records


def open_number_workbook(Workbook, load_workbook):
    """Load all existing group sheets and append without treating old data as done."""
    if NUMBER_WORKBOOK_PATH.exists():
        workbook = load_workbook(NUMBER_WORKBOOK_PATH)
        worksheet = workbook.worksheets[0]
        if migrate_legacy_group_blocks(workbook):
            from datetime import datetime
            import shutil
            backup = NUMBER_WORKBOOK_PATH.with_name(
                NUMBER_WORKBOOK_PATH.stem + '.before_group_append_' +
                datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.xlsx')
            shutil.copy2(NUMBER_WORKBOOK_PATH, backup)
            atomic_save(workbook, NUMBER_WORKBOOK_PATH)
            print(f'旧混合表已分群，原表保留为历史汇总；备份：{backup}')
    else:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "QQ号码"
        atomic_save(workbook, NUMBER_WORKBOOK_PATH)
    existing_numbers = set()
    for sheet in workbook.worksheets:
        headers = sheet_group_headers(sheet)
        if len(headers) == 1 and headers[0][0] == 1:
            existing_numbers.update(sheet_member_numbers(sheet))
    journal_records = load_journal_records()
    existing_numbers.update(str(record['number']).strip() for record in journal_records
                            if record['type'] == 'member')
    return workbook, worksheet, existing_numbers, set(), journal_records


def worksheet_row_has_data(worksheet, row):
    return any(cell.value is not None for cell in worksheet[row])


def find_group_section_in_sheet(worksheet, group_number):
    group_number = str(group_number).strip()
    header_rows = []
    for row in range(1, worksheet.max_row + 1):
        name = worksheet.cell(row=row, column=1).value
        value = worksheet.cell(row=row, column=2).value
        if (name is not None
                and GROUP_NUMBER_VALUE_PATTERN.fullmatch(str(value).strip())):
            header_rows.append((row, str(value).strip()))
    for index, (row, current_number) in enumerate(header_rows):
        if current_number != group_number:
            continue
        next_header_row = (
            header_rows[index + 1][0]
            if index + 1 < len(header_rows)
            else worksheet.max_row + 1
        )
        next_member_row = row + 1
        while next_member_row < next_header_row and worksheet_row_has_data(
                worksheet, next_member_row):
            next_member_row += 1
        return {
            "group_row": row,
            "next_member_row": next_member_row,
            "group_number": group_number,
        }
    return None


def append_group_info_to_sheet(worksheet, group_name, group_number):
    group_name = str(group_name or "").strip()
    group_number = str(group_number or "").strip()
    if not group_name:
        raise ValueError("群名称为空，不能保存群信息。")
    if not GROUP_NUMBER_VALUE_PATTERN.fullmatch(group_number):
        raise ValueError("群号格式无效，不能保存群信息。")

    existing_section = find_group_section_in_sheet(worksheet, group_number)
    if existing_section is not None:
        return existing_section

    last_used_row = max(
        (
            row
            for row in range(1, worksheet.max_row + 1)
            if worksheet_row_has_data(worksheet, row)
        ),
        default=0,
    )
    group_row = last_used_row + 3 if last_used_row else 1
    worksheet.cell(row=group_row, column=1).value = group_name
    number_cell = worksheet.cell(row=group_row, column=2)
    number_cell.number_format = "@"
    number_cell.value = group_number
    return {
        "group_row": group_row,
        "next_member_row": group_row + 1,
        "group_number": group_number,
    }


# 在工作表末尾追加一个群区块，并先记录群名称和群号。
def append_group_info(workbook, worksheet, group_name, group_number):
    group_name = str(group_name or '').strip()
    group_number = str(group_number or '').strip()
    if not group_name or not GROUP_NUMBER_VALUE_PATTERN.fullmatch(group_number):
        raise ValueError('群名称或群号无效，不能保存群信息。')
    sheet, created = ensure_group_sheet(workbook, worksheet, group_name, group_number)
    # Include group metadata in this file's own journal, even for old sheets.
    append_journal_record({'type': 'group', 'group_name': group_name, 'group_number': group_number})
    if created:
        atomic_save(workbook, NUMBER_WORKBOOK_PATH)
    return {'group_row': 1, 'next_member_row': sheet.max_row + 1,
            'group_number': group_number, 'group_sheet': sheet.title,
            'numbers': sheet_member_numbers(sheet)}


def append_number(workbook, worksheet, existing_numbers, number, group_section):
    """Deduplicate within this group; a shared member belongs in both groups."""
    sheet = workbook[group_section['group_sheet']]
    numbers = group_section.get('numbers')
    if numbers is None:
        numbers = group_section['numbers'] = sheet_member_numbers(sheet)
    if number in numbers:
        return False
    try:
        append_journal_record({'type': 'member', 'group_number': group_section['group_number'], 'number': number})
        sheet.cell(sheet.max_row + 1, 1, number).number_format = '@'
        numbers.add(number)
        existing_numbers.add(number)
        group_section['next_member_row'] = sheet.max_row + 1
    except Exception as error:
        raise CaptureStorageError(f'号码日志保存失败，停止采集：{error}') from error
    return True


def reconcile_journal(workbook, worksheet, journal_records, force_save=False):
    """Recover only this workbook's journal, using one worksheet per group."""
    changed = migrate_legacy_group_blocks(workbook)
    groups = {}
    numbers = {}
    for sheet in workbook.worksheets:
        headers = sheet_group_headers(sheet)
        if len(headers) == 1 and headers[0][0] == 1:
            key = headers[0][2]
            groups[key] = sheet
            numbers[key] = sheet_member_numbers(sheet)
    for record in journal_records:
        kind = record['type']
        key = str(record['group_number']).strip()
        if kind == 'group_complete':
            continue
        if kind == 'group':
            if key not in groups:
                groups[key], created = ensure_group_sheet(workbook, worksheet, record['group_name'], key)
                numbers[key] = sheet_member_numbers(groups[key])
                changed = created or changed
            continue
        if key not in groups:
            raise RuntimeError(f'号码日志中的群号 {key} 没有对应群信息。')
        number = str(record['number']).strip()
        if number not in numbers[key]:
            groups[key].cell(groups[key].max_row + 1, 1, number).number_format = '@'
            numbers[key].add(number)
            changed = True
    if changed or force_save:
        atomic_save(workbook, NUMBER_WORKBOOK_PATH)
    return changed or force_save


def collect_session_rows(session_list):  # 收集当前可见的会话行。
    """读取会话列表中当前可访问的会话名称，不查看消息或成员内容。"""
    panel = session_list.rectangle()
    rows = []
    positions = set()
    for row in session_list.descendants():
        if row.element_info.control_type not in ("Group", "ListItem", "DataItem"):
            continue
        rect = row.rectangle()
        # Unread badges are nested Groups too. Require a full-width, visible row.
        if (rect.width() < panel.width() * 0.6 or rect.height() <= 0
                or not panel.top <= (rect.top + rect.bottom) / 2 < panel.bottom
                or rect.left < panel.left or rect.right > panel.right + 2):
            continue
        texts = [child for child in row.children()
                 if child.element_info.control_type == "Text"
                 and child.window_text().strip()]
        if not texts:
            continue
        # The title is on the first line, left of the timestamp. Preserve numeric names.
        title = min(texts, key=lambda child: (child.rectangle().top, child.rectangle().left))
        position = (rect.top, rect.bottom)
        if position not in positions:
            positions.add(position)
            rows.append((title.window_text().strip(), row))
    return sorted(rows, key=lambda entry: (entry[1].rectangle().top, entry[1].rectangle().left))


def scroll_session_list(session_list, pyautogui, work_area):
    """推进群会话列表并保留重叠；优先使用 UIA，失败时使用小步长滚轮。"""
    rect = session_list.rectangle()
    if rect.width() <= 0 or rect.height() <= 0:
        return False, "会话列表区域没有有效屏幕范围"
    center_x = (rect.left + rect.right) // 2
    center_y = (rect.top + rect.bottom) // 2
    if not point_in_work_area(center_x, center_y, work_area):
        return False, "会话列表滚动目标位于系统工作区之外"

    try:
        scroll = session_list.iface_scroll
        if not scroll.CurrentVerticallyScrollable:
            return False, "群列表没有可滚动内容或已经到达底部"
        percent = float(scroll.CurrentVerticalScrollPercent)
        view_size = float(scroll.CurrentVerticalViewSize)
        if percent < 0 or percent >= 99.99 or not (0 < view_size < 100):
            return False, "群列表已经到达底部"
        scrollable_pixels = rect.height() * (100 / view_size - 1)
        if scrollable_pixels <= 0:
            return False, "群列表没有可滚动内容或已经到达底部"
        step_pixels = max(1.0, rect.height() * SESSION_SCROLL_STEP_RATIO)
        target = min(100.0, percent + step_pixels / scrollable_pixels * 100)
        if target <= percent:
            return False, "群列表已经到达底部"
        scroll.SetScrollPercent(-1.0, target)
        return True, "已使用 QQ 会话列表滚动接口推进，并保留页面重叠"
    except Exception as error:
        # 某些 QQ 版本不暴露 ScrollPattern；滚轮只推进小步，下一轮会重新读取可见群名。
        try:
            pyautogui.moveTo(center_x, center_y)
            pyautogui.scroll(-max(1, SESSION_SCROLL_WHEEL_NOTCHES), x=center_x, y=center_y)
            return True, f"QQ 未提供会话滚动接口，已使用小步长滚轮（{type(error).__name__}）"
        except Exception as wheel_error:
            return False, f"会话列表滚动失败：{wheel_error}"


def find_qq_windows(desktop, win32gui):
    """枚举可见的 QQ 主窗口，并按进程号去重。"""
    windows = []
    process_ids = set()
    for window in desktop.windows():
        try:
            if window.window_text().strip() != WINDOW_TITLE:
                continue
            process_id = window.process_id()
            handle = window.handle
            iconic = bool(win32gui.IsIconic(handle))
            if iconic:
                rect = win32gui.GetWindowPlacement(handle)[4]
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]
            else:
                rect = window.rectangle()
                width = rect.right - rect.left
                height = rect.bottom - rect.top
            if (
                width <= 0
                or height <= 0
                or process_id in process_ids
                or (not win32gui.IsWindowVisible(handle) and not iconic)
            ):
                continue
            process_ids.add(process_id)
            windows.append(window)
        except Exception as error:
            raise RuntimeError(
                "无法可靠检查一个 QQ 主窗口，已停止以避免选错账号。"
            ) from error
    return windows


def find_main_qq_window(desktop):  # 查找已按 QQ 号确认的主窗口。
    """只返回已确认进程中的 QQ 主窗口，不回退到任意窗口。"""
    if SELECTED_PROCESS_ID is None:
        return None
    for window in desktop.windows():
        try:
            if (window.window_text().strip() == WINDOW_TITLE
                    and window.process_id() == SELECTED_PROCESS_ID):
                return window
        except Exception:
            continue
    return None


def find_own_avatar(window):
    """只从账号栏定位本人头像，避免点击聊天或群成员头像。"""
    candidates = []
    account_header = []
    for control in window.descendants():
        try:
            name = control.window_text().strip()
            if (control.element_info.control_type != "Button"
                    or not control.is_visible() or not control.is_enabled()
                    or not (name == AVATAR_BUTTON_TEXT
                            or (AVATAR_AUTO_DETECT and name.endswith("的头像")))):
                continue
            candidates.append(control)
            if any(sibling.is_visible() and sibling.window_text().strip() == "在线状态"
                   for sibling in control.parent().children()):
                account_header.append(control)
        except Exception:
            continue
    if len(account_header) == 1:
        return account_header[0]
    if len(account_header) > 1:
        raise RuntimeError("检测到多个账号头像，无法确认本人账号栏。")
    rect = window.rectangle()
    header_candidates = []
    for control in candidates:
        area = control.rectangle()
        if (rect.left <= area.left < rect.left + min(180, (rect.right - rect.left) * 0.3)
                and rect.top <= area.top < rect.top + min(180, (rect.bottom - rect.top) * 0.25)):
            header_candidates.append(control)
    return header_candidates[0] if len(header_candidates) == 1 else None


def account_profile_roots(window, desktop):
    """Query only profile containers, without walking every chat/message control."""
    names = {'资料卡', '个人资料', '我的资料'}
    roots = []
    for top in desktop.windows(process=window.process_id()):
        try:
            if top.process_id() != window.process_id() or not top.is_visible():
                continue
            if top.window_text().strip() in names:
                roots.append(top)
                continue
            for name in names:
                for control in top.descendants(title=name):
                    if (control.element_info.control_type in {'Window', 'Pane', 'Group', 'Document'}
                            and control.is_visible()):
                        roots.append(control)
            for document in top.descendants(control_type='Document'):
                text = document.window_text().strip()
                if (document.is_visible() and all(marker in text for marker in
                                                 ('QQ空间', '编辑资料', '发消息'))):
                    roots.append(document)
        except Exception:
            continue
    return roots


def has_open_account_profile(window, desktop):
    """Detect a loading profile before clicking an avatar that would toggle it shut."""
    return bool(account_profile_roots(window, desktop))


def find_own_profile_account(window, desktop):
    """只从同一 QQ 进程的本人资料卡读取账号号码。"""
    profile_names = {"资料卡", "个人资料", "我的资料"}
    combined_pattern = re.compile(r"QQ\s*(?:号|号码)?\s*[:：]?\s*([0-9]{5,12})", re.IGNORECASE)
    label_pattern = re.compile(r"QQ\s*(?:号|号码)?\s*[:：]?", re.IGNORECASE)
    accounts = set()
    for scope in account_profile_roots(window, desktop):
        try:
            nodes = [scope, *scope.descendants()]
            scopes = [scope] if scope.window_text().strip() in profile_names else []
        except Exception:
            continue
        for scope in scopes:
            try:
                controls = [control for control in nodes[1:] if control.is_visible()]
                if not any(control.element_info.control_type == "Button"
                           and control.window_text().strip() in {"编辑资料", "修改资料"}
                           for control in controls):
                    continue
            except Exception:
                continue
            for control in controls:
                try:
                    if control.element_info.control_type not in {"Text", "Edit"}:
                        continue
                    name = control.window_text().strip()
                    match = combined_pattern.fullmatch(name)
                    if match:
                        accounts.add(match.group(1))
                        continue
                    if not label_pattern.fullmatch(name):
                        continue
                    label_rect = control.rectangle()
                    parent = control.parent()
                    if parent is None:
                        continue
                    for sibling in parent.children():
                        if (not sibling.is_visible()
                                or sibling.element_info.control_type not in {"Text", "Edit"}):
                            continue
                        number = sibling.window_text().strip()
                        if not re.fullmatch(r"[0-9]{5,12}", number):
                            continue
                        rect = sibling.rectangle()
                        overlap = min(label_rect.bottom, rect.bottom) - max(label_rect.top, rect.top)
                        height = min(label_rect.bottom - label_rect.top, rect.bottom - rect.top)
                        if (height > 0 and overlap >= height * 0.5
                                and -3 <= rect.left - label_rect.right <= max(80, height * 6)):
                            accounts.add(number)
                except Exception:
                    continue

        # Some QQ versions expose the own profile card only as one web Document.
        # Require its own-profile actions together with the QQ-space markers before
        # accepting a number from that combined text.
        for control in nodes:
            try:
                if (control.element_info.control_type != "Document"
                        or not control.is_visible()):
                    continue
                text = control.window_text().strip()
                if not all(marker in text for marker in ("QQ空间", "编辑资料", "发消息")):
                    continue
                accounts.update(combined_pattern.findall(text))
            except Exception:
                continue
    if len(accounts) > 1:
        raise RuntimeError("同一 QQ 进程检测到多个不同的本人 QQ 号，无法确认账号。")
    return next(iter(accounts), None)


def activate_window(window, win32con, win32gui, *, timeout=0.8, maximize=False):
    handle = window.handle
    if not win32gui.IsWindow(handle):
        raise RuntimeError('QQ 主窗口已失效，请重新选择已登录的 QQ。')
    if win32gui.IsIconic(handle) or not win32gui.IsWindowVisible(handle):
        win32gui.ShowWindow(handle, win32con.SW_RESTORE)
    if maximize:
        win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
    def ready():
        return (win32gui.IsWindow(handle) and win32gui.IsWindowVisible(handle)
                and win32gui.GetForegroundWindow() == handle
                and (not maximize or win32gui.GetWindowPlacement(handle)[1] == win32con.SW_MAXIMIZE))
    errors = []
    try:
        win32gui.BringWindowToTop(handle)
        win32gui.SetForegroundWindow(handle)
    except Exception as error:
        errors.append(str(error))
    if not ready():
        try:
            window.set_focus()
        except Exception as error:
            errors.append(str(error))
    if not after_ui_action(ready, timeout, interval=0.15):
        detail = '；'.join(errors)
        raise RuntimeError('无法激活所选 QQ 主窗口，请显示该窗口后重试。' + (f' 详情：{detail}' if detail else ''))


def read_qq_account_from_avatar(window, pyautogui, win32con, win32gui, desktop):
    """复用已打开的本人资料卡；必要时点击本人头像并等待号码出现。"""
    account = find_own_profile_account(window, desktop)
    if account is not None:
        return account
    # The card can already be open while its accessibility tree is still loading.
    # Do not click the avatar again until that existing card has been read.
    if has_open_account_profile(window, desktop):
        print('资料卡已打开，正在等待本人 QQ 号码显示。')
        account = wait_for_result(
            lambda: find_own_profile_account(window, desktop),
            AVATAR_TIMEOUT_SECONDS, interval=0.2,
        )
        if account is not None:
            return account
        raise RuntimeError(
            '已打开的资料卡未能确认本人 QQ 号。请关闭资料卡，'
            '点击主窗口左上角自己的头像，确认显示“编辑资料”和 QQ 号码后重试。'
        )
    activate_window(window, win32con, win32gui)
    avatar_deadline = time.monotonic() + 2.0
    avatar = None
    while avatar is None and time.monotonic() < avatar_deadline:
        avatar = find_own_avatar(window)
        if avatar is None:
            time.sleep(0.2)
    if avatar is None:
        raise RuntimeError("未找到 QQ 主窗口账号栏中的本人头像。")
    print('正在打开本人资料卡，等待 QQ 号码显示。')
    avatar.click_input()
    deadline = time.monotonic() + AVATAR_TIMEOUT_SECONDS
    retry_at = time.monotonic() + AVATAR_TIMEOUT_SECONDS / 2
    retried = False
    while time.monotonic() < deadline:
        check_stop()
        account = find_own_profile_account(window, desktop)
        if account is not None:
            return account
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        # A foreground transition can swallow the first click. Retry only when
        # no card is open, once, within the original timeout budget.
        if not retried and time.monotonic() >= retry_at:
            retried = True
            if not has_open_account_profile(window, desktop):
                print('未检测到资料卡，正在重新打开本人头像。')
                activate_window(window, win32con, win32gui)
                avatar = find_own_avatar(window)
                if avatar is not None and time.monotonic() < deadline:
                    avatar.click_input()
        remaining = deadline - time.monotonic()
        if remaining > 0:
            time.sleep(min(0.2, remaining))
    if not has_open_account_profile(window, desktop):
        raise RuntimeError(
            '已点击本人头像，但 QQ 没有打开资料卡。请先手动点击头像；'
            '若也没有反应，请重启这个 QQ 并重新登录。确认资料卡能显示“编辑资料”和 QQ 号码后再开始。'
        )
    raise RuntimeError(
        '点击本人头像后，未能从资料卡中确认 QQ 号。请点击主窗口左上角自己的头像，'
        '确认资料卡显示“编辑资料”和 QQ 号码，再重新开始。'
    )


def select_qq_window_by_number(desktop, pyautogui, win32con, win32gui):
    """逐个确认候选窗口的登录 QQ 号，并要求目标号唯一匹配。"""
    windows = find_qq_windows(desktop, win32gui)
    if not windows:
        raise RuntimeError("没有检测到可用的 QQ 主窗口。请先登录 QQ 并显示主窗口。")

    matches = []
    for window in windows:
        process_id = window.process_id()
        try:
            account = read_qq_account_from_avatar(
                window, pyautogui, win32con, win32gui, desktop
            )
        except Exception as error:
            raise RuntimeError(
                f"无法可靠确认 QQ 进程 {process_id} 的登录号码，已停止：{error}"
            ) from error
        if account is None:
            raise RuntimeError(
                f"无法读取 QQ 进程 {process_id} 的登录号码，已停止；不会猜测窗口。"
            )
        print(f"已识别 QQ 号 {account}（进程 {process_id}）。")
        if account == TARGET_QQ_NUMBER.strip():
            matches.append(window)

    if not matches:
        raise RuntimeError(f"没有找到登录 QQ 号为 {TARGET_QQ_NUMBER.strip()} 的窗口。")
    if len(matches) > 1:
        raise RuntimeError(
            f"检测到多个窗口登录 QQ 号 {TARGET_QQ_NUMBER.strip()}，无法唯一选择，已停止。"
        )
    return matches[0]


def get_session_list(main_window):  # 查找会话列表控件。
    """从 QQ 主窗口的 UI Automation 树中找到会话列表。"""
    return next(
        (
            control
            for control in main_window.descendants()
            if control.window_text().strip() == SESSION_LIST_TEXT
        ),
        None,
    )


def get_contact_group_list(main_window):
    return next(
        (control for control in main_window.descendants(control_type='Pane')
         if control.element_info.control_type == "Pane"
         and control.window_text().strip() == "联系人列表"
         and control.rectangle().width() > 0
         and control.rectangle().height() > 0),
        None,
    )


def get_joined_groups_category(main_window):
    matches = [
        control for control in main_window.descendants(control_type='Button')
        if control.element_info.control_type == "Button"
        and control.window_text().strip().startswith("我加入的群聊")
        and control.is_visible()
        and control.rectangle().width() > 0
        and control.rectangle().height() > 0
    ]
    return matches[0] if len(matches) == 1 else None


def collect_joined_group_rows(main_window, work_area):
    """Collect visible group rows under Contacts > Groups > Joined groups."""
    pane = get_contact_group_list(main_window)
    category = get_joined_groups_category(main_window)
    if pane is None or category is None:
        return []
    pane_rect = pane.rectangle()
    excluded_prefixes = (
        "置顶群聊", "未命名的群聊", "我创建的群聊", "我管理的群聊", "我加入的群聊",
    )
    rows = []
    positions = set()
    for control in pane.descendants(control_type='Button'):
        try:
            if control.element_info.control_type != "Button" or not control.is_visible():
                continue
            name = control.window_text().strip()
            if not name or name.startswith(excluded_prefixes):
                continue
            rect = control.rectangle()
            center_x = (rect.left + rect.right) // 2
            center_y = (rect.top + rect.bottom) // 2
            if (rect.width() <= 0 or rect.height() <= 0
                    or rect.left < pane_rect.left or rect.right > pane_rect.right + 2
                    or not point_in_work_area(center_x, center_y, work_area)):
                continue
            position = (rect.top, rect.bottom)
            if position in positions:
                continue
            positions.add(position)
            rows.append((name, control))
        except Exception:
            continue
    return sorted(rows, key=lambda entry: entry[1].rectangle().top)


def ensure_joined_groups_page(desktop, work_area):
    """Navigate to Contacts > Groups and expand the joined-groups category."""
    main_window = find_main_qq_window(desktop)
    if main_window is None:
        raise RuntimeError("无法定位已确认的 QQ 主窗口")
    if get_contact_group_list(main_window) is None:
        contacts = [
            control for control in main_window.descendants()
            if control.element_info.control_type == "Button"
            and control.window_text().strip() == "联系人"
            and control.is_visible()
        ]
        if len(contacts) != 1 or not click_control_in_work_area(contacts[0], work_area):
            raise RuntimeError("无法唯一定位联系人入口")
        deadline = time.monotonic() + max(3.0, WAIT_SECONDS * 3)
        while time.monotonic() < deadline:
            main_window = find_main_qq_window(desktop)
            if main_window and get_contact_group_list(main_window):
                break
            time.sleep(0.15)

    group_tabs = [
        control for control in main_window.descendants()
        if control.element_info.control_type == "TabItem"
        and control.window_text().strip() == "群聊"
        and control.is_visible()
    ]
    if len(group_tabs) != 1:
        raise RuntimeError("联系人页中无法唯一定位“群聊”标签")
    try:
        selected = group_tabs[0].is_selected()
    except Exception:
        selected = None
    if selected is not True:
        group_tabs[0].click_input()
        def joined_category_ready():
            current = find_main_qq_window(desktop)
            return get_joined_groups_category(current) if current else None
        if not after_ui_action(joined_category_ready, min(0.5, WAIT_SECONDS), interval=0.15):
            raise RuntimeError("群聊页中无法定位“我加入的群聊”分类")

    main_window = find_main_qq_window(desktop)
    category = get_joined_groups_category(main_window) if main_window else None
    if category is None:
        raise RuntimeError("群聊页中无法定位“我加入的群聊”分类")
    if not collect_joined_group_rows(main_window, work_area):
        if not click_control_in_work_area(category, work_area):
            raise RuntimeError("“我加入的群聊”分类不在可操作区域")
        deadline = time.monotonic() + max(3.0, WAIT_SECONDS * 3)
        while time.monotonic() < deadline:
            main_window = find_main_qq_window(desktop)
            if main_window and collect_joined_group_rows(main_window, work_area):
                break
            time.sleep(0.15)
    if main_window is None or get_contact_group_list(main_window) is None:
        raise RuntimeError("未能打开联系人群聊列表")
    return main_window


def scroll_joined_group_list(desktop, main_window, pyautogui, work_area):
    last_error = ""
    prefer_wheel = False
    for attempt in range(3):
        main_window = find_main_qq_window(desktop)
        pane = get_contact_group_list(main_window) if main_window else None
        rows = collect_joined_group_rows(main_window, work_area) if main_window else []
        if pane is None or not rows:
            return False, "我加入的群聊列表没有可滚动的可见群行"
        rect = pane.rectangle()
        left, top = max(rect.left, work_area[0]), max(rect.top, work_area[1])
        right, bottom = min(rect.right, work_area[2]), min(rect.bottom, work_area[3])
        if right - left < 12 or bottom - top < 12:
            return False, "我加入的群聊列表不在可操作屏幕区域"
        baseline = tuple(
            (name, row.rectangle().top, row.rectangle().bottom)
            for name, row in rows
        )
        method = ""
        if not prefer_wheel:
            try:
                scroll = pane.iface_scroll
                if not scroll.CurrentVerticallyScrollable:
                    return False, "我加入的群聊列表已到末尾"
                percent = float(scroll.CurrentVerticalScrollPercent)
                view_size = float(scroll.CurrentVerticalViewSize)
                if percent >= 99.99:
                    return False, "我加入的群聊列表已到末尾"
                if not (0 <= percent < 100 and 0 < view_size < 100):
                    raise ValueError("无效的滚动范围")
                scrollable_pixels = rect.height() * (100 / view_size - 1)
                if scrollable_pixels <= 0:
                    return False, "我加入的群聊列表没有可滚动内容"
                step = min((bottom - top) * 0.65, scrollable_pixels)
                target = min(100.0, percent + step / scrollable_pixels * 100)
                if target <= percent:
                    return False, "我加入的群聊列表已到末尾"
                scroll.SetScrollPercent(-1.0, target)
                method = "群聊列表滚动接口"
            except Exception as error:
                last_error = type(error).__name__
                prefer_wheel = True
        if not method:
            row_index = min(len(rows) - 1, ((attempt + 1) * len(rows)) // 4)
            row_rect = rows[row_index][1].rectangle()
            x = max(left, min(right - 1, (row_rect.left + row_rect.right) // 2))
            y = max(top, min(bottom - 1, (row_rect.top + row_rect.bottom) // 2))
            try:
                pyautogui.moveTo(x, y)
                pyautogui.scroll(-1, x=x, y=y)
                method = "群行内滚轮"
            except Exception as error:
                return False, f"我加入的群聊列表滚动失败：{error}"

        deadline = time.monotonic() + max(1.0, WAIT_SECONDS)
        if not CONTINUE_WHEN_READY:
            time.sleep(0.15)
        while time.monotonic() < deadline:
            current_window = find_main_qq_window(desktop)
            current_rows = collect_joined_group_rows(current_window, work_area) if current_window else []
            signature = tuple(
                (name, row.rectangle().top, row.rectangle().bottom)
                for name, row in current_rows
            )
            if signature and signature != baseline:
                return True, f"{method}已推进，可见群行已变化"
            time.sleep(0.15)
        prefer_wheel = True
        print(f"群聊列表暂未变化，滚动尝试 {attempt + 1}/3。")

    detail = f"（列表接口提示：{last_error}）" if last_error else ""
    return False, f"连续 3 次滚动后可见群行未变化，可能已到末尾或滚动未生效{detail}"


def enter_group_chat_from_profile(main_window, desktop, group_name, work_area):
    send_buttons = [
        control for control in main_window.descendants()
        if control.element_info.control_type == "Button"
        and control.window_text().strip() in {"发消息", "发送消息", "发送群聊消息"}
        and control.is_visible()
    ]
    if len(send_buttons) != 1 or not click_control_in_work_area(send_buttons[0], work_area):
        raise RuntimeError(f"群“{group_name}”资料页中无法唯一定位发消息按钮")
    return wait_for_group_header(desktop, group_name)


def open_group_profile_from_joined_list(desktop, group_name, pyautogui, work_area):
    main_window = ensure_joined_groups_page(desktop, work_area)
    main_window = reset_joined_groups_to_top(desktop, main_window, pyautogui, work_area)

    for _ in range(min(MAX_SCROLLS, 100)):
        main_window = find_main_qq_window(desktop)
        rows = collect_joined_group_rows(main_window, work_area) if main_window else []
        matches = [row for name, row in rows if name == group_name]
        if len(matches) == 1:
            if not click_control_in_work_area(matches[0], work_area):
                raise RuntimeError(f"群聊“{group_name}”不在可操作区域")
            deadline = time.monotonic() + max(5.0, WAIT_SECONDS * 5)
            while time.monotonic() < deadline:
                main_window = find_main_qq_window(desktop)
                details = find_group_profile_details(main_window, group_name) if main_window else None
                if details is not None:
                    print(f"已在群资料页核对群名和群号：{group_name}（{details[0]}）。")
                    return main_window, details[0]
                time.sleep(0.2)
            raise RuntimeError(f"点击群聊“{group_name}”后，资料页未显示可确认的名称和群号")
        if len(matches) > 1:
            raise RuntimeError(f"“我加入的群聊”列表中出现多个同名群“{group_name}”，已停止以免选错")

        scrolled, reason = scroll_joined_group_list(
            desktop, main_window, pyautogui, work_area
        )
        if not scrolled:
            raise RuntimeError(f"未能在“我加入的群聊”列表中定位“{group_name}”：{reason}")
        if not CONTINUE_WHEN_READY:
            time.sleep(min(0.5, WAIT_SECONDS))
    raise RuntimeError(f"搜索“我加入的群聊”列表达到滚动上限，未找到“{group_name}”")


def ensure_message_session_list(desktop, work_area):
    main_window = find_main_qq_window(desktop)
    if main_window is None:
        raise RuntimeError("无法定位已确认的 QQ 主窗口")
    if get_session_list(main_window) is not None:
        return main_window
    messages = [
        control for control in main_window.descendants()
        if control.element_info.control_type == "Button"
        and control.window_text().strip() == "消息"
        and control.is_visible()
    ]
    if len(messages) != 1 or not click_control_in_work_area(messages[0], work_area):
        raise RuntimeError("无法唯一定位左侧“消息”入口")
    deadline = time.monotonic() + max(3.0, WAIT_SECONDS * 3)
    while time.monotonic() < deadline:
        main_window = find_main_qq_window(desktop)
        if main_window and get_session_list(main_window) is not None:
            return main_window
        time.sleep(0.15)
    raise RuntimeError("点击“消息”后仍无法定位会话列表")


def get_group_header_button(main_window, group_name):  # 查找指定群聊标题按钮。
    """通过群名和成员人数特征寻找群聊标题按钮，不访问成员列表内容。"""
    buttons = main_window.descendants(control_type="Button")
    for control in buttons:
        titles = [control.window_text().strip()]
        child_texts = [child.window_text().strip()
                       for child in control.children()
                       if child.element_info.control_type == "Text" and child.window_text().strip()]
        if child_texts:
            titles.append(" ".join(child_texts))
        for title in titles:
            # Normalize full-width parentheses without changing user-supplied patterns.
            normalized = title.replace("（", "(").replace("）", ")")
            for candidate in (title, normalized):
                match = GROUP_HEADER_PATTERN.search(candidate)
                expected = group_name if candidate == title else group_name.replace("（", "(").replace("）", ")")
                if match and candidate[:match.start()].rstrip() == expected:
                    return control
        # Current QQ can omit the count. Only accept an exact title alongside
        # group-only toolbar actions; an ordinary contact name alone is insufficient.
        if group_name in titles:
            rect = control.rectangle()
            if rect.width() <= 0 or rect.height() <= 0:
                continue
            for action in buttons:
                if action.window_text().strip() not in ("群应用", "邀请加群"):
                    continue
                action_rect = action.rectangle()
                if (action_rect.width() > 0 and action_rect.height() > 0
                        and action_rect.left >= rect.right
                        and action_rect.top < rect.bottom and action_rect.bottom > rect.top):
                    return control
    return None


def find_group_number_in_profile(main_window, group_name):
    """从群资料页的名称、纯数字群号和成员数同一行关系中取群号。"""
    page_details = find_group_profile_details(main_window, group_name)
    if page_details is not None:
        return page_details[0]

    candidates = set()
    for document in main_window.descendants():
        try:
            if (document.element_info.control_type != "Document"
                    or not document.is_visible()):
                continue
            controls = [
                control for control in document.descendants()
                if control.is_visible()
                and control.element_info.control_type == "Text"
                and control.rectangle().width() > 0
                and control.rectangle().height() > 0
            ]
        except Exception:
            continue

        profile_markers = {control.window_text().strip() for control in controls}
        if not profile_markers.intersection({"群介绍", "群简介", "群分类", "群资料"}):
            continue
        names = [control for control in controls
                 if control.window_text().strip() == group_name]
        numbers = [control for control in controls
                   if GROUP_NUMBER_VALUE_PATTERN.fullmatch(control.window_text().strip())]
        member_counts = [control for control in controls
                         if GROUP_MEMBER_LABEL_PATTERN.fullmatch(control.window_text().strip())]
        for name_control in names:
            name_rect = name_control.rectangle()
            for number_control in numbers:
                number_rect = number_control.rectangle()
                vertical_gap = number_rect.top - name_rect.bottom
                if (not 0 <= vertical_gap <= 28
                        or abs(number_rect.left - name_rect.left) > 36):
                    continue
                for count_control in member_counts:
                    count_rect = count_control.rectangle()
                    overlap = min(number_rect.bottom, count_rect.bottom) - max(
                        number_rect.top, count_rect.top
                    )
                    row_height = min(number_rect.height(), count_rect.height())
                    horizontal_gap = count_rect.left - number_rect.right
                    if (row_height > 0 and overlap >= row_height * 0.5
                            and -6 <= horizontal_gap <= 120):
                        candidates.add(number_control.window_text().strip())

    if len(candidates) > 1:
        raise RuntimeError("群资料页出现多个符合条件的群号，无法确认，已停止。")
    return next(iter(candidates), None)


def find_group_profile_details(main_window, group_name):
    """Read identity from the Contacts group profile page shown in current QQ."""
    texts = []
    buttons = []
    for control in main_window.descendants():
        try:
            if not control.is_visible():
                continue
            rect = control.rectangle()
            if rect.width() <= 0 or rect.height() <= 0:
                continue
            kind = control.element_info.control_type
            if kind == "Text":
                texts.append((control.window_text().strip(), rect))
            elif kind == "Button" and control.window_text().strip() in {
                "发消息", "发送消息", "发送群聊消息",
            }:
                buttons.append(control)
        except Exception:
            continue

    names = [rect for text, rect in texts if text == group_name]
    labels = [rect for text, rect in texts if text == "群号"]
    numbers = [(text, rect) for text, rect in texts
               if GROUP_NUMBER_VALUE_PATTERN.fullmatch(text)]
    group_member_markers = [text for text, _ in texts
                            if text.startswith("群成员") and re.search(r"\d+\s*人", text)]
    candidates = set()
    for name_rect in names:
        for label_rect in labels:
            name_to_label_gap = label_rect.top - name_rect.bottom
            if (not 0 <= name_to_label_gap <= 24
                    or abs(label_rect.left - name_rect.left) > 80):
                continue
            for number, number_rect in numbers:
                overlap = min(label_rect.bottom, number_rect.bottom) - max(
                    label_rect.top, number_rect.top
                )
                row_height = min(label_rect.height(), number_rect.height())
                horizontal_gap = number_rect.left - label_rect.right
                if (row_height > 0 and overlap >= row_height * 0.5
                        and -4 <= horizontal_gap <= 80):
                    candidates.add(number)
    if not names or not group_member_markers or len(buttons) != 1:
        return None
    if len(candidates) > 1:
        raise RuntimeError(f"群资料页中“{group_name}”对应多个群号，无法确认。")
    if len(candidates) == 1:
        return next(iter(candidates)), buttons[0]
    return None


def profile_card_number_snapshot(main_window, desktop):
    """Read numeric IDs exposed by separate QQ profile-card windows."""
    process_id = main_window.process_id()
    cards = {}
    for window in desktop.windows():
        try:
            if (window.handle == main_window.handle
                    or window.process_id() != process_id
                    or not window.is_visible()
                    or window.window_text().strip() not in {"资料卡", "群资料", "群聊资料"}):
                continue
            numbers = {
                control.window_text().strip()
                for control in window.descendants()
                if control.is_visible()
                and control.element_info.control_type in {"Text", "Edit"}
                and GROUP_NUMBER_VALUE_PATTERN.fullmatch(control.window_text().strip())
                and control.rectangle().width() > 0
                and control.rectangle().height() > 0
            }
            cards[window.handle] = numbers
        except Exception:
            continue
    return cards


def find_new_group_number_in_profile_card(main_window, desktop, previous_cards):
    """Accept a unique number only from a profile card opened or changed by the title click."""
    current_cards = profile_card_number_snapshot(main_window, desktop)
    candidates = set()
    for handle, numbers in current_cards.items():
        if handle not in previous_cards or numbers != previous_cards[handle]:
            if len(numbers) == 1:
                candidates.update(numbers)
    if len(candidates) > 1:
        raise RuntimeError("群标题打开的资料卡包含多个号码，无法确认群号。")
    return next(iter(candidates), None)


def read_group_number_from_profile(main_window, group_name, desktop=None):
    """必要时点击群标题进入群资料页，并等待明确的群号字段出现。"""
    group_number = find_group_number_in_profile(main_window, group_name)
    if group_number is not None:
        return group_number

    previous_profile_cards = (
        profile_card_number_snapshot(main_window, desktop) if desktop is not None else {}
    )
    header_button = get_group_header_button(main_window, group_name)
    if header_button is None:
        raise RuntimeError(f"无法定位群聊“{group_name}”的标题，不能读取群号。")
    name_controls = [
        control for control in header_button.children()
        if control.element_info.control_type == "Text"
        and control.window_text().strip() == group_name
    ]
    click_target = name_controls[0] if len(name_controls) == 1 else header_button
    if click_target is header_button:
        print(f"群名文本控件不唯一，改用已确认的群标题按钮打开“{group_name}”资料。")
    click_target.click_input()
    deadline = time.monotonic() + 5.0
    while time.monotonic() < deadline:
        group_number = find_group_number_in_profile(main_window, group_name)
        if group_number is not None:
            return group_number
        if desktop is not None:
            group_number = find_new_group_number_in_profile_card(
                main_window, desktop, previous_profile_cards
            )
            if group_number is not None:
                print(f"从群标题打开的独立资料卡读取到群号 {group_number}。")
                return group_number
        time.sleep(0.2)
    raise RuntimeError(f"打开群“{group_name}”的资料页后，仍无法可靠识别 QQ 群号。")


def wait_for_group_header(desktop, group_name):
    """Wait for the chat to load instead of classifying an empty page as private chat."""
    deadline = time.monotonic() + max(3.0, WAIT_SECONDS * 3)
    main_window = None
    while True:
        main_window = find_main_qq_window(desktop)
        header = get_group_header_button(main_window, group_name) if main_window else None
        if header is not None or time.monotonic() >= deadline:
            return main_window, header
        time.sleep(0.25)


def ensure_group_chat_open(desktop, group_name, work_area):
    """Reuse an already-open group; click a session only when switching to it."""
    main_window = find_main_qq_window(desktop)
    header = get_group_header_button(main_window, group_name) if main_window else None
    if header is not None:
        print(f"目标群“{group_name}”已经打开，直接使用当前聊天，不重复点击会话。")
        return main_window, header
    session_list = get_session_list(main_window) if main_window else None
    row = next((row for name, row in collect_session_rows(session_list)
                if name == group_name), None) if session_list is not None else None
    if row is None:
        raise RuntimeError(f"无法重新定位聊天“{group_name}”")
    # Recheck after enumerating rows in case the selected chat finished loading.
    main_window = find_main_qq_window(desktop)
    header = get_group_header_button(main_window, group_name) if main_window else None
    if header is not None:
        print(f"目标群“{group_name}”已经打开，直接使用当前聊天，不重复点击会话。")
        return main_window, header
    if not click_control_in_work_area(row, work_area):
        raise RuntimeError(f"会话“{group_name}”不在系统工作区内")
    if not CONTINUE_WHEN_READY:
        time.sleep(WAIT_SECONDS)
    return wait_for_group_header(desktop, group_name)


def member_panel_is_open(main_window):  # 判断成员列表面板是否已打开。
    """只检查成员列表面板是否存在，不读取面板中的任何成员信息。"""
    return get_member_list_window(main_window) is not None


def get_member_list_window(main_window):  # 查找群成员列表窗口。
    names = {MEMBER_LIST_TEXT, "成员列表", "群成员列表"}
    return next(
        (
            control
            for name in names for control in main_window.descendants(title=name)
            if control.window_text().strip() in names
            and control.element_info.control_type in ("Pane", "Window", "Group", "List")
            and control.rectangle().width() > 0 and control.rectangle().height() > 0
        ),
        None,
    )


def get_visible_member_panel(main_window):  # 查找成员列表的可见外框。
    """QQ 的 UIA Window 矩形有时是虚拟滚动区域，使用可见 Group 外框定位滚动条。"""
    for member_list in (control for name in {MEMBER_LIST_TEXT, '成员列表', '群成员列表'}
                        for control in main_window.descendants(title=name)):
        if (member_list.window_text().strip() in (MEMBER_LIST_TEXT, "成员列表", "群成员列表")
                and member_list.element_info.control_type in ("Pane", "List")
                and member_list.rectangle().width() > 0 and member_list.rectangle().height() > 0):
            return member_list
    titled_panel = next(
        (
            control
            for control in main_window.descendants(control_type="Group")
            if control.window_text().strip() in (MEMBER_LIST_TEXT, "成员列表", "群成员列表")
            and control.rectangle().width() > 0
            and control.rectangle().height() > 0
        ),
        None,
    )
    if titled_panel is not None:
        return titled_panel

    # 部分 QQ 版本不给外框设置名称；此时选择右侧最大的可见分组作为面板。
    main_rect = main_window.rectangle()
    candidates = []
    for control in main_window.descendants(control_type="Group"):
        rect = control.rectangle()
        if (
            rect.width() >= 120
            and rect.height() >= 200
            and rect.right >= main_rect.right - 24
            and rect.top >= main_rect.top
            and rect.bottom <= main_rect.bottom
        ):
            candidates.append((rect.width() * rect.height(), control))
    return max(candidates, default=(0, None), key=lambda entry: entry[0])[1]


def collect_member_items(member_window, visible_panel=None, work_area=None):  # 收集成员列表中的成员项。
    """优先读取 UIA 列表项；部分 QQ 版本会把成员行暴露为分组控件。"""
    panel = (visible_panel if visible_panel is not None else member_window).rectangle()
    left, top, right, bottom = panel.left, panel.top, panel.right, panel.bottom
    if work_area is not None:
        left, top = max(left, work_area[0]), max(top, work_area[1])
        right, bottom = min(right, work_area[2]), min(bottom, work_area[3])
    for control_type in ("ListItem", "DataItem", "Group"):
        items = []
        positions = set()
        for item in member_window.descendants(control_type=control_type):
            rect = item.rectangle()
            is_inside_panel = (
                rect.width() >= (right - left) * 0.5
                and rect.height() > 0
                and rect.left < right
                and rect.right > left
                and top <= (rect.top + rect.bottom) / 2 < bottom
            )
            position = (rect.top, rect.bottom)
            if is_inside_panel and position not in positions:
                if control_type == "Group" and not any(
                        child.element_info.control_type == "Text" and child.window_text().strip()
                        for child in item.children()):
                    continue
                positions.add(position)
                items.append(item)
        if items:
            return sorted(items, key=lambda item: (item.rectangle().top, item.rectangle().left))
    return []


def member_item_text(item):  # 获取成员行及其可访问子文本。
    texts = [item.window_text().strip()]
    texts.extend(
        child.window_text().strip()
        for child in item.descendants()
        if child.window_text().strip()
    )
    return " ".join(texts)


# 检查成员行及其子控件中是否存在完全匹配的指定名称。
def member_item_has_exact_name(item, expected_name):
    texts = [item.window_text().strip()]
    texts.extend(
        child.window_text().strip()
        for child in item.descendants()
        if child.window_text().strip()
    )
    return expected_name in texts


# 从 QQ 界面可访问文本中读取群成员总数，供滚动距离估算使用。
def get_member_count(main_window):
    for control in main_window.descendants(control_type="Text"):
        match = MEMBER_COUNT_PATTERN.search(control.window_text().strip())
        if match:
            return int(match.group(1))
    return 0


def member_item_identity(item):
    text = member_item_text(item)
    try:
        runtime_id = item.element_info.runtime_id
    except Exception:
        runtime_id = None
    if runtime_id:
        try:
            runtime_id = tuple(runtime_id)
        except TypeError:
            runtime_id = (runtime_id,)
    return text, runtime_id


def member_page_overlap(previous_keys, current_keys):
    """Count rows repeated at the end of one visible page and start of the next."""
    limit = min(len(previous_keys), len(current_keys))
    previous_text_counts = {}
    current_text_counts = {}
    for text, _ in previous_keys:
        previous_text_counts[text] = previous_text_counts.get(text, 0) + 1
    for text, _ in current_keys:
        current_text_counts[text] = current_text_counts.get(text, 0) + 1

    for count in range(limit, 0, -1):
        previous_page_tail = previous_keys[-count:]
        current_page_head = current_keys[:count]
        text_matches = all(
            previous[0]
            and previous[0] == current[0]
            and previous_text_counts[previous[0]] == 1
            and current_text_counts[current[0]] == 1
            for previous, current in zip(previous_page_tail, current_page_head)
        )
        identities_match = all(
            previous == current for previous, current in zip(previous_page_tail, current_page_head)
        )
        if identities_match or (count >= 2 and text_matches):
            return count
    return 0


def member_page_signature(items, panel=None):
    return tuple(member_item_text(item) for item in items)


def scroll_member_list(desktop, pyautogui, work_area):
    """Advance the member list with UIA, falling back to a wheel inside the panel."""
    main_window = find_main_qq_window(desktop)
    member_list = get_member_list_window(main_window) if main_window else None
    panel = get_visible_member_panel(main_window) if main_window else None
    if member_list is None or panel is None:
        raise RuntimeError("翻页前无法定位群成员列表")
    items = collect_member_items(member_list, panel, work_area)
    baseline = member_page_signature(items, panel)
    if not baseline:
        return False, "成员列表没有可见条目，未执行滚动"
    last_error = ""
    prefer_wheel = False
    for attempt in range(3):
        main_window = find_main_qq_window(desktop)
        member_list = get_member_list_window(main_window) if main_window else None
        panel = get_visible_member_panel(main_window) if main_window else None
        if member_list is None or panel is None:
            raise RuntimeError("翻页时群成员列表已关闭")
        rect = panel.rectangle()
        left, top = max(rect.left, work_area[0]), max(rect.top, work_area[1])
        right, bottom = min(rect.right, work_area[2]), min(rect.bottom, work_area[3])
        if right - left < 12 or bottom - top < 12:
            return False, "成员列表不在可操作的屏幕区域内"
        method = None
        if not prefer_wheel:
            try:
                scroll = panel.iface_scroll
                if not scroll.CurrentVerticallyScrollable:
                    return False, "成员列表已全部显示，无需继续向下滚动"
                percent = float(scroll.CurrentVerticalScrollPercent)
                view_size = float(scroll.CurrentVerticalViewSize)
                if percent >= 99.99:
                    return False, "成员列表已到最底部"
                if not (0 <= percent < 100 and 0 < view_size < 100):
                    raise ValueError("无效的滚动范围")
                # Keep overlap between pages, even when fewer than 13 rows fit.
                row_height = statistics.median(item.rectangle().height() for item in items)
                step = min(row_height * MEMBER_PAGE_SIZE, (bottom - top) * 0.7)
                scrollable_pixels = rect.height() * (100 / view_size - 1)
                target = min(100.0, percent + step / scrollable_pixels * 100)
                scroll.SetScrollPercent(-1.0, target)
                method = "成员列表滚动接口"
            except Exception as error:
                last_error = type(error).__name__
                prefer_wheel = True
        if method is None:
            # Wheel over a visible member row; gaps and panel padding may ignore it.
            wheel_items = collect_member_items(member_list, panel, work_area)
            if not wheel_items:
                return False, "成员列表没有可滚动的可见成员行"
            row_index = min(
                len(wheel_items) - 1,
                ((attempt + 1) * len(wheel_items)) // 4,
            )
            row_rect = wheel_items[row_index].rectangle()
            x = max(left, min(right - 1, (row_rect.left + row_rect.right) // 2))
            y = max(top, min(bottom - 1, (row_rect.top + row_rect.bottom) // 2))
            pyautogui.moveTo(x, y)
            pyautogui.scroll(-1, x=x, y=y)
            method = "成员行内滚轮"
        deadline = time.monotonic() + max(1.0, WAIT_SECONDS)
        if not CONTINUE_WHEN_READY:
            time.sleep(0.15)
        while True:
            main_window = find_main_qq_window(desktop)
            member_list = get_member_list_window(main_window) if main_window else None
            panel = get_visible_member_panel(main_window) if main_window else None
            if member_list is not None and panel is not None:
                current = collect_member_items(member_list, panel, work_area)
                signature = member_page_signature(current, panel)
                if signature and signature != baseline:
                    return True, f"{method}已推进，当前可见 {len(current)} 名成员"
            if time.monotonic() >= deadline:
                break
            time.sleep(0.15)
        prefer_wheel = True
        print(f"成员列表暂未变化，翻页尝试 {attempt + 1}/3。")
    detail = f"（接口提示：{last_error}）" if last_error else ""
    return False, f"连续 3 次滚动后成员列表未变化，可能已到底或滚动被资料卡遮挡{detail}"


# 按页处理普通群成员；点击后先识别号码，成功保存后再处理下一位。
def click_non_admin_members(
    main_window,
    pyautogui,
    desktop,
    work_area,
    win32gui,
    win32con,
    ocr_state,
    number_workbook,
    number_worksheet,
    existing_numbers,
    group_section,
):  # 逐页点击符合条件的群成员并保存可识别的号码。
    member_window = get_member_list_window(main_window)
    if member_window is None:
        raise RuntimeError("成员列表窗口不可访问")

    clicked_count = 0
    skipped_member_count = 0
    page_turns = 0
    saved_number_count = 0
    duplicate_number_count = 0
    number_error_count = 0
    previous_page_keys = ()
    resume = group_section.get('resume', {})
    resume_index, page_turns = restore_capture_position(main_window, desktop, pyautogui, work_area, resume)
    initial_saved_count = max(resume.get('saved_count', 0), group_section.get('saved_count', 0))
    group_section['completed'] = False
    if initial_saved_count >= MAX_MEMBERS_PER_GROUP:
        group_section['completed'] = True
        return (0, 0, page_turns, 0, 0, 0)
    while True:
        check_pause()
        main_window = find_main_qq_window(desktop)
        member_window = get_member_list_window(main_window) if main_window else None
        if member_window is None:
            raise RuntimeError("翻页后无法重新定位成员列表")

        visible_panel = get_visible_member_panel(main_window)
        if visible_panel is None:
            raise RuntimeError("无法定位成员列表的可见区域")
        items = collect_member_items(member_window, visible_panel, work_area)
        page_items = [(member_item_text(item), item) for item in items]
        if not page_items:
            # An empty UIA result is not proof that all members were collected.
            print('成员列表为空，保留未完成进度以便下次继续。')
            return (
                clicked_count,
                skipped_member_count,
                page_turns,
                saved_number_count,
                duplicate_number_count,
                number_error_count,
            )

        page_keys = [member_item_identity(item) for _, item in page_items]
        overlap = member_page_overlap(previous_page_keys, page_keys)
        if overlap:
            print(f"跳过与上一屏重叠的成员行：{overlap} 名。")
        previous_page_keys = page_keys
        page_names = [key[0] for key in page_keys]
        first_index = max(overlap, resume_index)
        resume_index = 0
        save_member_position(group_section, main_window, work_area, page_names, first_index,
                             page_turns, initial_saved_count + saved_number_count)

        for member_index, (role_and_name, item) in enumerate(page_items[first_index:], start=first_index):
            check_pause()
            if not role_and_name:
                continue
            if (
                "群主" in role_and_name
                or "管理员" in role_and_name
                or any(
                    member_item_has_exact_name(item, name)
                    for name in EXCLUDED_MEMBER_NAMES
                )
            ):
                skipped_member_count += 1
                continue
            qq_process_id = main_window.process_id() if main_window else -1
            qq_main_handle = main_window.handle if main_window else 0
            stage_started = time.monotonic()
            try:
                before_snapshot = collect_qq_snapshot(
                    desktop,
                    qq_process_id,
                    qq_main_handle,
                    win32gui,
                    win32con,
                )
            except Exception as error:
                before_snapshot = {"controls": {}, "windows": {}}
                number_error_count += 1
                print(f"无法读取点击前的 UIA 快照：{type(error).__name__}")
            report_slow_member_stage('点击前读取', stage_started)
            stage_started = time.monotonic()
            rect = item.rectangle()
            clicked_rect = (rect.left, rect.top, rect.right, rect.bottom)
            try:
                if not click_control_in_work_area(item, work_area, visible_panel.rectangle(), avatar=True):
                    number_error_count += 1
                    print("成员控件完全位于系统工作区外，已安全跳过该成员。")
                    continue
            except Exception as error:
                number_error_count += 1
                print(f"点击成员失败，继续处理其他成员：{type(error).__name__}")
                continue
            report_slow_member_stage('点击成员', stage_started)
            clicked_count += 1
            if not CONTINUE_WHEN_READY:
                time.sleep(MEMBER_CLICK_DELAY)
            stage_started = time.monotonic()
            try:
                # 保持当前资料卡不变，在最多三轮 UIA/OCR 尝试后再继续成员循环。
                number = None
                last_diagnostic = "尚未尝试识别"
                after_snapshot = before_snapshot
                for attempt in range(1, MAX_NUMBER_RECOGNITION_ATTEMPTS + 1):
                    if attempt > 1:
                        print(
                            f"资料卡号码未识别，正在第 {attempt}/"
                            f"{MAX_NUMBER_RECOGNITION_ATTEMPTS} 次重试。"
                        )
                        time.sleep(PROFILE_POLL_INTERVAL)
                    try:
                        # UIA 仅接受点击后变化的号码控件，避免误用前一位成员的旧号码。
                        number, ui_diagnostic, after_snapshot = extract_member_number_from_popup(
                            desktop,
                            before_snapshot,
                            qq_process_id,
                            qq_main_handle,
                            clicked_rect,
                            win32gui,
                            win32con,
                        )
                        if number is None:
                            number, ocr_diagnostic = read_number_with_ocr(
                                pyautogui,
                                ocr_state,
                                before_snapshot,
                                after_snapshot,
                                clicked_rect,
                                work_area,
                                qq_main_handle,
                            )
                            last_diagnostic = (
                                f"UIA: {ui_diagnostic}; OCR: {ocr_diagnostic}"
                            )
                    except Exception as error:
                        last_diagnostic = f"{type(error).__name__}: {error}"
                    if number is not None:
                        break
                # 达到尝试上限仍无号码时记录一次失败，避免无限等待。
                if number is None:
                    raise RuntimeError(
                        f"连续 {MAX_NUMBER_RECOGNITION_ATTEMPTS} 次未确认号码；"
                        f"{last_diagnostic}"
                    )
                report_slow_member_stage('等待并读取号码', stage_started)
                stage_started = time.monotonic()
                if append_number(
                    number_workbook,
                    number_worksheet,
                    existing_numbers,
                    number,
                    group_section,
                ):
                    saved_number_count += 1
                    print(f"已追加保存号码：{number}")
                    if initial_saved_count + saved_number_count >= MAX_MEMBERS_PER_GROUP:
                        save_member_position(group_section, main_window, work_area, page_names, member_index + 1,
                                             page_turns, initial_saved_count + saved_number_count)
                        group_section['completed'] = number_error_count == 0
                        print(
                            f"本群累计保存 {initial_saved_count + saved_number_count} 名成员，"
                            f"达到上限 {MAX_MEMBERS_PER_GROUP}，继续下一个群。"
                        )
                        return (clicked_count, skipped_member_count, page_turns,
                                saved_number_count, duplicate_number_count,
                                number_error_count)
                else:
                    duplicate_number_count += 1
                    print(f"号码已存在，跳过重复项：{number}")
            except Exception as error:
                number_error_count += 1
                print(f"读取或保存成员号码失败，继续处理：{error}")
            save_member_position(group_section, main_window, work_area, page_names, member_index + 1,
                                 page_turns, initial_saved_count + saved_number_count)
            report_slow_member_stage('号码和进度保存', stage_started)

        if page_turns >= MAX_MEMBER_PAGE_TURNS:
            group_section['completed'] = number_error_count == 0
            print(f"成员采集达到向下翻页上限 {MAX_MEMBER_PAGE_TURNS}。")
            return (clicked_count, skipped_member_count, page_turns, saved_number_count,
                    duplicate_number_count, number_error_count)
        advanced, reason = scroll_member_list(desktop, pyautogui, work_area)
        print(f"成员翻页：{reason}。")
        if not advanced:
            group_section['completed'] = number_error_count == 0 and reason in {
                '成员列表已全部显示，无需继续向下滚动', '成员列表已到最底部', '已到底'}
            return (clicked_count, skipped_member_count, page_turns, saved_number_count,
                    duplicate_number_count, number_error_count)
        page_turns += 1
        print(f"成员列表已向下翻页 {page_turns}/{MAX_MEMBER_PAGE_TURNS}，继续采集。")


# 初始化依赖和工作簿，再按名额扫描会话并累计成员处理统计。
def main():  # 执行 QQ 群聊面板查看主流程。
    """依次检查最多三个非排除聊天，并处理其中的群成员面板。"""
    global SELECTED_PROCESS_ID, _capture_workbook, _capture_worksheet, _capture_progress
    _capture_workbook = _capture_worksheet = _capture_progress = None

    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="backslashreplace")

    print("QQ 群成员面板查看助手")
    print("脚本确认群名称和群号后打开成员列表，按规则点击普通成员并保存群信息和新识别号码。")
    print(f"当前系统：{platform.platform()}")
    print('等待模式：完成后立即继续（原超时和重试上限不变）。' if CONTINUE_WHEN_READY
          else '等待模式：按已保存的固定等待时间执行。')
    if sys.platform != "win32":
        print("此脚本只支持 Windows QQ 桌面版。")
        return
    target_qq_number = TARGET_QQ_NUMBER.strip()
    if not target_qq_number:
        print("尚未填写目标 QQ 号。请修改脚本顶部的 TARGET_QQ_NUMBER 后重试。")
        return
    if not re.fullmatch(r"\d{5,12}", target_qq_number):
        print("TARGET_QQ_NUMBER 格式无效，请填写 5 到 12 位数字后重试。")
        return

    (
        Desktop,
        pyautogui,
        win32api,
        win32con,
        win32gui,
        Workbook,
        load_workbook,
        RapidOCR,
    ) = load_dependencies()
    if RapidOCR is None:
        print("OCR 后备未启用；需要时运行：python -m pip install rapidocr-onnxruntime")
    ocr_state = {"factory": RapidOCR, "engine": None}

    pyautogui.FAILSAFE = True
    print("兼容性提示：需要 QQ 桌面版已登录，主窗口可见且 UI Automation 可访问。")
    desktop = Desktop(backend="uia")
    try:
        main_window = select_qq_window_by_number(
            desktop, pyautogui, win32con, win32gui
        )
        SELECTED_PROCESS_ID = main_window.process_id()
        print(f"已选择目标 QQ 号 {target_qq_number}（进程 {SELECTED_PROCESS_ID}）。")
    except Exception as error:
        print(f"账号选择失败：{error}")
        raise

    try:
        (
            number_workbook,
            number_worksheet,
            existing_numbers,
            existing_group_numbers,
            journal_records,
        ) = open_number_workbook(Workbook, load_workbook)
        reconcile_journal(number_workbook, number_worksheet, journal_records)
        _capture_workbook, _capture_worksheet = number_workbook, number_worksheet
        _capture_progress = CaptureProgress(NUMBER_WORKBOOK_PATH, target_qq_number)
        existing_group_numbers = completed_groups_for_resume(journal_records, _capture_progress)
        set_pause_flush(flush_capture)
    except Exception as error:
        print(f"无法打开或创建号码工作簿 {NUMBER_WORKBOOK_PATH}：{error}")
        raise
    initial_numbers = set(existing_numbers)
    print(f"号码将保存到：{NUMBER_WORKBOOK_PATH}")
    print(f"号码追加日志：{NUMBER_JOURNAL_PATH}")
    print(f"采集进度：{_capture_progress.path}；Ctrl+B 暂停/继续。")
    print(f'已读取本账号的完成记录：{len(existing_group_numbers)} 个群；核对群号后跳过，未完成群继续采集。')
    check_pause()
    print("新版 QQ 的界面或缩放比例变化可能导致控件无法识别；无法识别时会自动结束或跳过。")

    main_window = find_main_qq_window(desktop)
    if main_window is None:
        print(f"找不到已确认的 QQ 进程 {SELECTED_PROCESS_ID}，已停止。")
        return

    # 窗口准备代码块：恢复并最大化 QQ，让会话列表控件可以被 UI Automation 发现。
    try:
        activate_window(main_window, win32con, win32gui, timeout=WAIT_SECONDS, maximize=True)
    except Exception as error:
        print(f"无法激活或最大化 QQ 窗口：{error}")
        raise

    try:
        work_area = get_work_area(win32api, win32con, main_window.handle)
        print(f"鼠标工作区：{work_area}（任务栏区域已排除）")
    except Exception as error:
        print(f"无法确定 Windows 工作区边界，为避开任务栏已停止：{error}")
        return

    try:
        main_window = ensure_joined_groups_page(desktop, work_area)
    except Exception as error:
        print(f"无法打开“我加入的群聊”列表：{error}")
        raise
    try:
        main_window = reset_joined_groups_to_top(desktop, main_window, pyautogui, work_area)
    except Exception as error:
        print(f"无法确认群聊列表位于顶部：{error}")
        raise

    # Process visible joined groups from top to bottom, then advance the group list.
    seen_names = set()
    pending_retry_names = set()
    session_retry_counts = {}
    session_retry_limit = max(0, SESSION_RETRY_LIMIT)
    checked_group_count = 0
    excluded_chat_count = 0
    non_group_count = 0
    opened_count = 0
    failed_count = 0
    member_clicked_count = 0
    admin_skipped_count = 0
    member_page_turn_count = 0
    member_number_saved_count = 0
    duplicate_number_count = 0
    member_number_error_count = 0
    existing_group_skipped_count = 0
    scroll_count = 0

    list_complete = False
    while scroll_count <= MAX_SCROLLS:
        check_pause()
        main_window = find_main_qq_window(desktop)
        if main_window is None or get_contact_group_list(main_window) is None:
            failed_count += 1
            print("无法重新定位“我加入的群聊”列表，结束扫描。")
            break

        new_rows = [
            (name, row)
            for name, row in collect_joined_group_rows(main_window, work_area)
            if name not in seen_names or name in pending_retry_names
        ]
        if not new_rows:
            scrolled, scroll_reason = scroll_joined_group_list(
                desktop, main_window, pyautogui, work_area
            )
            if not scrolled:
                if "已到末尾" in scroll_reason or "没有可滚动内容" in scroll_reason:
                    list_complete = True
                    print(f"我加入的群聊列表已处理到底部：{scroll_reason}。")
                else:
                    failed_count += 1
                    print(f"群聊列表无法继续推进，扫描停止：{scroll_reason}。")
                break
            print(f"群聊列表推进：{scroll_reason}。")
            if not CONTINUE_WHEN_READY:
                time.sleep(WAIT_SECONDS)
            scroll_count += 1
            continue

        for name, _row in new_rows:
            check_pause()
            # Apply the saved exclusions before opening a profile or counting a group.
            if name in EXCLUDED_CHAT_NAMES:
                seen_names.add(name)
                pending_retry_names.discard(name)
                session_retry_counts.pop(name, None)
                excluded_chat_count += 1
                print(f"群聊“{name}”在排除名单中，已跳过。")
                continue
            if name not in pending_retry_names:
                checked_group_count += 1
            print(f"进度：按列表顺序检查第 {checked_group_count} 个群聊：{name}")
            try:
                main_window, group_number = open_group_profile_from_joined_list(
                    desktop, name, pyautogui, work_area
                )
            except Exception as error:
                retry_count = session_retry_counts.get(name, 0) + 1
                session_retry_counts[name] = retry_count
                if retry_count <= session_retry_limit:
                    pending_retry_names.add(name)
                    print(
                        f"群聊“{name}”在我加入的群聊列表中暂时无法定位；"
                        f"将在下一轮重新查找（第 {retry_count}/{session_retry_limit} 次重试）。"
                    )
                    try:
                        ensure_joined_groups_page(desktop, work_area)
                    except Exception as restore_error:
                        print(f"返回群聊列表失败：{restore_error}")
                    continue
                pending_retry_names.discard(name)
                seen_names.add(name)
                failed_count += 1
                print(
                    f"群聊“{name}”在我加入的群聊列表中初次查找及 {session_retry_limit} 次重试后仍未定位，"
                    f"已停止重试并记录失败：{error}"
                )
                try:
                    ensure_joined_groups_page(desktop, work_area)
                except Exception as restore_error:
                    print(f"返回群聊列表失败：{restore_error}")
                continue

            try:
                if group_number in existing_group_numbers:
                    existing_group_skipped_count += 1
                    print(f"群号 {group_number} 已完成采集，跳过群“{name}”的成员处理。")
                    continue

                main_window, header_button = enter_group_chat_from_profile(
                    main_window, desktop, name, work_area
                )
                if header_button is None:
                    raise RuntimeError(f"点击群资料页的发消息按钮后，未能确认进入群聊“{name}”")

                # Some QQ layouts already show the member list in the right sidebar.
                if not member_panel_is_open(main_window):
                    if not click_control_in_work_area(header_button, work_area):
                        failed_count += 1
                        print(f"群聊“{name}”标题不在系统工作区内，已安全跳过。")
                        continue
                    main_window = wait_for_member_panel(desktop, work_area)
                else:
                    print(f"使用“{name}”右侧已显示的群成员列表。")
                    if CONTINUE_WHEN_READY:
                        main_window = wait_for_member_panel(desktop, work_area)
                if not main_window or not member_panel_is_open(main_window):
                    failed_count += 1
                    print(f"无法确认“{name}”的成员列表已打开，跳过该群并继续列表顺序。")
                    continue

                print(f"已按资料页核对并进入群聊：{name}（群号 {group_number}）")
                group_section = append_group_info(
                    number_workbook,
                    number_worksheet,
                    name,
                    group_number,
                )
                group_section.update(progress=_capture_progress, group_name=name,
                                     resume=resume_for_new_task(_capture_progress.get(group_number)),
                                     saved_count=0)
                opened_count += 1
                print(f"已保存群信息：{name}（群号 {group_number}）")
                try:
                    (
                        clicked_count,
                        skipped_member_count,
                        page_turns,
                        saved_number_count,
                        duplicate_count,
                        number_error_count,
                    ) = click_non_admin_members(
                        main_window,
                        pyautogui,
                        desktop,
                        work_area,
                        win32gui,
                        win32con,
                        ocr_state,
                        number_workbook,
                        number_worksheet,
                        existing_numbers,
                        group_section,
                    )
                except Exception as error:
                    failed_count += 1
                    print(f"点击群成员时发生错误：{error}")
                    continue
                try:
                    if group_section.get('completed'):
                        append_journal_record({
                            "type": "group_complete",
                            "group_number": group_number,
                            "account": target_qq_number,
                        })
                        sync_journal()
                        _capture_progress.save(group_number, group_name=name, completed=True)
                        existing_group_numbers.add(group_number)
                    else:
                        print(f'群“{name}”尚未确认完成，已保留续采进度。')
                    flush_capture()
                except Exception as error:
                    failed_count += 1
                    print(f"群“{name}”完成标记保存失败，脚本停止：{error}")
                    return
                member_clicked_count += clicked_count
                admin_skipped_count += skipped_member_count
                member_page_turn_count += page_turns
                member_number_saved_count += saved_number_count
                duplicate_number_count += duplicate_count
                member_number_error_count += number_error_count
                print(
                    f"本群点击普通成员：{clicked_count}；跳过群主、管理员或指定成员：{skipped_member_count}；"
                    f"翻页：{page_turns}/{MAX_MEMBER_PAGE_TURNS}"
                )
                if page_turns >= MAX_MEMBER_PAGE_TURNS:
                    print("已达到本群翻页上限，继续检查下一个聊天。")
                else:
                    print("当前群聊成员处理已结束，继续检查下一个聊天。")
            except Exception as error:
                failed_count += 1
                print(f"处理群聊“{name}”失败，继续列表中的下一个群：{error}")
            finally:
                seen_names.add(name)
                pending_retry_names.discard(name)
                session_retry_counts.pop(name, None)
                try:
                    ensure_joined_groups_page(desktop, work_area)
                except Exception as restore_error:
                    failed_count += 1
                    print(f"处理完群“{name}”后未能返回“我加入的群聊”列表：{restore_error}")

    if pending_retry_names:
        pending = "、".join(sorted(pending_retry_names))
        print(f"仍有群聊未能定位：{pending}")
    elif list_complete:
        print(f"结束原因：已按顺序处理“我加入的群聊”列表中的 {checked_group_count} 个群聊。")
    elif scroll_count > MAX_SCROLLS:
        failed_count += 1
        print(f"结束原因：已达到群聊列表滚动次数上限 {MAX_SCROLLS}，列表可能未处理完整。")
    if opened_count == 0:
        print("本次未打开任何群成员列表，因此未采集号码；请检查群列表与资料页识别提示。")
    try:
        sync_journal()
        if reconcile_journal(
            number_workbook,
            number_worksheet,
            load_journal_records(),
            force_save=True,
        ):
            print("已将追加日志一次性合并到号码工作簿。")
    except Exception as error:
        failed_count += 1
        print(f"追加日志合并到号码工作簿失败：{error}；日志仍已保留，可下次启动恢复。")
    print_summary(
        opened_count,
        non_group_count,
        failed_count,
        excluded_chat_count,
        checked_group_count,
    )
    print(f"因群号已记录而跳过的群数：{existing_group_skipped_count}")
    print(f"点击普通成员总数：{member_clicked_count}")
    print(f"跳过群主、管理员或指定成员总数：{admin_skipped_count}")
    print(f"成员列表翻页总数：{member_page_turn_count}")
    print(f"新保存号码总数：{member_number_saved_count}")
    print(f"跳过重复号码总数：{duplicate_number_count}")
    print(f"号码读取或保存失败总数：{member_number_error_count}")


# 打印已检查聊天、异常和跳过项目的汇总统计。
def print_summary(
    opened_count,
    skipped_count,
    failed_count,
    excluded_count=0,
    checked_count=0,
):  # 输出处理统计汇总。
    """输出脚本处理情况。"""
    print("\n处理汇总")
    print(f"成功打开成员列表的群聊：{opened_count}")
    print(f"未能打开成员列表或识别群资料的群聊：{skipped_count}")
    print(f"跳过的指定群聊：{excluded_count}")
    print(f"按列表顺序检查的群聊：{checked_count}")
    print(f"自动处理失败数：{failed_count}")
    print("成员名称和角色仅用于识别及跳过成员；工作簿按群保存群名称、群号和全局去重的成员号码。")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n收到中断信号，脚本已安全停止。")
