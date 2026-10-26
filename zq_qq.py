"""自动检查最多三个 QQ 聊天并按规则逐页点击群成员。"""  # 模块用途说明。

from __future__ import annotations  # 启用延迟类型注解求值。

import platform  # 读取当前操作系统信息。

import statistics
import re  # 提供群标题正则匹配功能。
import sys  # 读取 Python 运行环境信息。
import time  # 提供等待和超时计时功能。
from pathlib import Path  # 定位号码工作簿。
from task_control import atomic_save


WAIT_SECONDS = 1.0  # 设置界面操作之间的等待时间。
MEMBER_CLICK_DELAY = 1.0  # 每次点击群成员后的等待时间。
PROFILE_READ_TIMEOUT = 2.0  # 等待资料卡文本稳定的最长时间。
PROFILE_POLL_INTERVAL = 0.25  # 资料卡 UIA 轮询间隔。
MAX_NUMBER_RECOGNITION_ATTEMPTS = 3  # 每位成员号码识别的最多尝试次数。
MEMBER_PAGE_SIZE = 14  # 每次翻页目标推进的成员行数。
MEMBER_TRIGGER_POSITION = 13  # 兼容旧配置；自动分页已不依赖此门槛。
MAX_MEMBER_PAGE_TURNS = 2  # 每个群聊最多向下翻页次数，可按需调整。
MAX_SCROLLS = 500  # 限制会话列表最大滚动次数。
MAX_CHATS = 3  # 最多自动检查前三个不同的非排除聊天窗口。
MAX_GROUPS = 2  # 最多打开并处理两个实际群聊的成员列表。
NO_PROGRESS_LIMIT = 3  # 限制连续无新会话时的结束阈值。
GROUP_HEADER_PATTERN = re.compile(r"\(\s*\d+\s*\)\s*$")  # 匹配群标题末尾的成员数量。
MEMBER_COUNT_PATTERN = re.compile(r"群聊成员\s*(\d+)")  # 读取成员面板显示的总人数。
EXCLUDED_CHAT_NAMES = {
    "蒂玖mod交流群",
    "CodebotAPI客服2️⃣群",
    "AI资源共享群",
    "tokkenShop中转站",
}  # 精确匹配并跳过的群聊名称，不计入检查名额。
EXCLUDED_MEMBER_NAMES = {"Q群管家"}  # 按成员显示名称精确匹配的跳过名单。
NUMBER_WORKBOOK_PATH = Path(__file__).resolve().with_name("qq_member_numbers.xlsx")
NUMBER_TEXT_PATTERN = re.compile(r"^\s*(\d{5,12})\s*$")
WINDOW_TITLE = "QQ"
SESSION_LIST_TEXT = "会话列表"
MEMBER_LIST_TEXT = "成员列表"
TARGET_PROCESS_ID = 0
USE_OCR = True


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
def click_control_in_work_area(control, work_area, clip_rect=None):
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
    x = (left + right) // 2
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
def collect_qq_snapshot(desktop, process_id, main_handle, win32gui, win32con):
    """Capture accessible controls and popup bounds for the QQ process and its owned windows."""
    snapshot = {"controls": {}, "windows": {}}
    for window in desktop.windows():
        try:
            if not window.is_visible() or not is_qq_related_window(
                window, process_id, main_handle, win32gui, win32con
            ):
                continue
            handle = window.handle
            rect = window.rectangle()
            window_bounds = (rect.left, rect.top, rect.right, rect.bottom)
            title = window.window_text().strip()
            snapshot["windows"][handle] = {
                "title": title,
                "bounds": window_bounds,
                "process_id": window.process_id(),
            }
            controls = [window]
            controls.extend(window.descendants())
        except Exception:
            continue

        for control in controls:
            try:
                info = control.element_info
                text = (control.window_text() or "").strip()
                if not text:
                    text = (getattr(info, "name", "") or "").strip()
                control_type = getattr(info, "control_type", "Unknown") or "Unknown"
                rect = control.rectangle()
                if rect.width() <= 0 or rect.height() <= 0:
                    continue
                bounds = (rect.left, rect.top, rect.right, rect.bottom)
                signature = (handle, control_type, text, bounds)
                snapshot["controls"][signature] = {
                    "window_handle": handle,
                    "window_title": title,
                    "control_type": control_type,
                    "text": text,
                    "bounds": bounds,
                }
            except Exception:
                continue
    return snapshot


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
        latest_snapshot = collect_qq_snapshot(
            desktop, process_id, main_handle, win32gui, win32con
        )
        nearby_controls = []
        click_x = (clicked_rect[0] + clicked_rect[2]) // 2
        click_y = (clicked_rect[1] + clicked_rect[3]) // 2
        for control in latest_snapshot["controls"].values():
            left, top, right, bottom = control["bounds"]
            if (
                right >= click_x - 520
                and left <= click_x + 520
                and bottom >= click_y - 420
                and top <= click_y + 420
            ):
                nearby_controls.append(control)
        nearby_changed_controls = [
            control
            for control in new_snapshot_controls(before_snapshot, latest_snapshot)
            if control in nearby_controls
        ]
        number, _ = select_number_below_label(
            nearby_controls, candidate_controls=nearby_changed_controls
        )
        latest_diagnostic = snapshot_diagnostic(before_snapshot, latest_snapshot)
        if number is not None:
            return number, latest_diagnostic, latest_snapshot
        time.sleep(PROFILE_POLL_INTERVAL)
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


def open_number_workbook(Workbook, load_workbook):
    """打开或创建号码工作簿；已有号码按文本读取以支持去重。"""
    if NUMBER_WORKBOOK_PATH.exists():
        workbook = load_workbook(NUMBER_WORKBOOK_PATH)
        worksheet = workbook.active
    else:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "QQ号码"
        atomic_save(workbook, NUMBER_WORKBOOK_PATH)
    existing_numbers = set()
    for (value,) in worksheet.iter_rows(min_col=1, max_col=1, values_only=True):
        if value is not None:
            text = str(value).strip()
            if text.isdigit():
                existing_numbers.add(text)
    return workbook, worksheet, existing_numbers


# 将尚未记录的号码以文本格式追加到工作簿，并立即保存。
def append_number(workbook, worksheet, existing_numbers, number):
    if number in existing_numbers:
        return False
    populated_rows = [
        row
        for row in range(1, worksheet.max_row + 1)
        if worksheet.cell(row=row, column=1).value is not None
    ]
    row = max(populated_rows, default=0) + 1
    cell = worksheet.cell(row=row, column=1)
    cell.number_format = "@"
    cell.value = number
    atomic_save(workbook, NUMBER_WORKBOOK_PATH)
    existing_numbers.add(number)
    return True


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


def find_main_qq_window(desktop):  # 查找 QQ 主窗口。
    """查找已登录的 QQ 主窗口。"""
    windows = [window for window in desktop.windows()
               if window.window_text().strip() == WINDOW_TITLE
               and (not TARGET_PROCESS_ID or window.process_id() == TARGET_PROCESS_ID)]
    if not windows:
        return None
    return windows[0]


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
            for control in main_window.descendants()
            if control.window_text().strip() in names
            and control.element_info.control_type in ("Pane", "Window", "Group", "List")
            and control.rectangle().width() > 0 and control.rectangle().height() > 0
        ),
        None,
    )


def get_visible_member_panel(main_window):  # 查找成员列表的可见外框。
    """QQ 的 UIA Window 矩形有时是虚拟滚动区域，使用可见 Group 外框定位滚动条。"""
    for member_list in main_window.descendants():
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


def member_page_signature(items, panel):
    top = panel.rectangle().top
    return tuple((member_item_text(item), item.rectangle().top - top, item.rectangle().height())
                 for item in items)


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
            # A single notch keeps overlap; never wheel over the chat/session list.
            x = int(right - min(12, (right - left) / 4))
            y = int((top + bottom) / 2)
            pyautogui.moveTo(x, y)
            pyautogui.scroll(-1, x=x, y=y)
            method = "成员列表内滚轮"
        deadline = time.monotonic() + max(1.0, WAIT_SECONDS)
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
    while True:
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
            return (
                clicked_count,
                skipped_member_count,
                page_turns,
                saved_number_count,
                duplicate_number_count,
                number_error_count,
            )

        for role_and_name, item in page_items:
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
            rect = item.rectangle()
            clicked_rect = (rect.left, rect.top, rect.right, rect.bottom)
            try:
                if not click_control_in_work_area(item, work_area, visible_panel.rectangle()):
                    number_error_count += 1
                    print("成员控件完全位于系统工作区外，已安全跳过该成员。")
                    continue
            except Exception as error:
                number_error_count += 1
                print(f"点击成员失败，继续处理其他成员：{type(error).__name__}")
                continue
            clicked_count += 1
            time.sleep(MEMBER_CLICK_DELAY)
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
                if append_number(
                    number_workbook,
                    number_worksheet,
                    existing_numbers,
                    number,
                ):
                    saved_number_count += 1
                    print(f"已追加保存号码：{number}")
                else:
                    duplicate_number_count += 1
                    print(f"号码已存在，跳过重复项：{number}")
            except Exception as error:
                number_error_count += 1
                print(f"读取或保存成员号码失败，继续处理：{error}")

        if page_turns >= MAX_MEMBER_PAGE_TURNS:
            print(f"成员采集达到向下翻页上限 {MAX_MEMBER_PAGE_TURNS}。")
            return (clicked_count, skipped_member_count, page_turns, saved_number_count,
                    duplicate_number_count, number_error_count)
        advanced, reason = scroll_member_list(desktop, pyautogui, work_area)
        print(f"成员翻页：{reason}。")
        if not advanced:
            return (clicked_count, skipped_member_count, page_turns, saved_number_count,
                    duplicate_number_count, number_error_count)
        page_turns += 1
        print(f"成员列表已向下翻页 {page_turns}/{MAX_MEMBER_PAGE_TURNS}，继续采集。")


# 初始化依赖和工作簿，再按名额扫描会话并累计成员处理统计。
def main():  # 执行 QQ 群聊面板查看主流程。
    """依次检查最多三个非排除聊天，并处理其中的群成员面板。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="backslashreplace")

    print("QQ 群成员面板查看助手")
    print("脚本打开群成员列表，识别成员名称和角色后按规则点击，并仅保存新识别的数字号码。")
    print(f"当前系统：{platform.platform()}")
    if sys.platform != "win32":
        print("此脚本只支持 Windows QQ 桌面版。")
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
    try:
        number_workbook, number_worksheet, existing_numbers = open_number_workbook(
            Workbook, load_workbook
        )
    except Exception as error:
        print(f"无法打开或创建号码工作簿 {NUMBER_WORKBOOK_PATH}：{error}")
        return
    initial_numbers = set(existing_numbers)
    print(f"号码将保存到：{NUMBER_WORKBOOK_PATH}")
    pyautogui.FAILSAFE = True
    print("兼容性提示：需要 QQ 桌面版已登录，主窗口可见且 UI Automation 可访问。")
    print("新版 QQ 的界面或缩放比例变化可能导致控件无法识别；无法识别时会自动结束或跳过。")

    desktop = Desktop(backend="uia")
    main_window = find_main_qq_window(desktop)
    if main_window is None:
        print("未找到标题为 QQ 的桌面窗口。请打开并登录 QQ 后重新运行。")
        return

    # 窗口准备代码块：恢复并最大化 QQ，让会话列表控件可以被 UI Automation 发现。
    try:
        win32gui.ShowWindow(main_window.handle, win32con.SW_RESTORE)
        win32gui.ShowWindow(main_window.handle, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(main_window.handle)
        time.sleep(WAIT_SECONDS)
    except Exception as error:
        print(f"无法激活或最大化 QQ 窗口：{error}")
        return

    try:
        work_area = get_work_area(win32api, win32con, main_window.handle)
        print(f"鼠标工作区：{work_area}（任务栏区域已排除）")
    except Exception as error:
        print(f"无法确定 Windows 工作区边界，为避开任务栏已停止：{error}")
        return

    session_list = get_session_list(main_window)
    if session_list is None:
        print("QQ 没有暴露可访问的会话列表控件；请确认使用桌面版并重试。")
        return

    # 会话扫描代码块：只使用会话行的首个文本控件作为名称；通过滚动会话面板
    # 加载更多条目。滚动位置由会话列表窗口矩形计算，不访问群成员列表。
    seen_names = set()
    checked_chat_count = 0
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
    scroll_count = 0
    no_progress = 0

    while (
        checked_chat_count < MAX_CHATS
        and opened_count < MAX_GROUPS
        and scroll_count <= MAX_SCROLLS
        and no_progress < NO_PROGRESS_LIMIT
    ):
        main_window = find_main_qq_window(desktop)
        session_list = get_session_list(main_window) if main_window else None
        if session_list is None:
            failed_count += 1
            print("无法重新定位 QQ 会话列表，结束扫描。")
            break

        new_rows = [
            (name, row)
            for name, row in collect_session_rows(session_list)
            if name not in seen_names
        ]
        if not new_rows:
            no_progress += 1
        else:
            no_progress = 0

        for name, row in new_rows:
            if checked_chat_count >= MAX_CHATS or opened_count >= MAX_GROUPS:
                break
            seen_names.add(name)
            if name in EXCLUDED_CHAT_NAMES:
                excluded_chat_count += 1
                print(f"按完整名称跳过指定群聊：{name}")
                continue

            checked_chat_count += 1
            print(f"进度：正在检查第 {checked_chat_count}/{MAX_CHATS} 个非排除聊天：{name}")
            try:
                main_window, header_button = ensure_group_chat_open(desktop, name, work_area)
                if header_button is None:
                    non_group_count += 1
                    print(f"未能识别“{name}”的群聊标题：可能是私聊、聊天未打开或 QQ 界面不兼容；未进入成员采集。")
                    continue

                # Some QQ layouts already show the member list in the right sidebar.
                if not member_panel_is_open(main_window):
                    if not click_control_in_work_area(header_button, work_area):
                        failed_count += 1
                        print(f"群聊“{name}”标题不在系统工作区内，已安全跳过。")
                        continue
                    time.sleep(WAIT_SECONDS)
                    main_window = find_main_qq_window(desktop)
                else:
                    print(f"使用“{name}”右侧已显示的群成员列表。")
                if not main_window or not member_panel_is_open(main_window):
                    failed_count += 1
                    print(f"无法确认“{name}”的成员列表已打开，结束扫描。")
                    print_summary(
                        opened_count,
                        non_group_count,
                        failed_count,
                        excluded_chat_count,
                        checked_chat_count,
                    )
                    return

                opened_count += 1
                print(f"已打开群成员面板：{name}（{opened_count} 个群）")
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
                    )
                except Exception as error:
                    failed_count += 1
                    print(f"点击群成员时发生错误：{error}")
                    continue
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
                print(f"处理会话“{name}”失败，自动检查下一个聊天：{error}")

        if checked_chat_count >= MAX_CHATS or opened_count >= MAX_GROUPS:
            break

        # 将鼠标放在会话列表区域再滚动，避免滚动到群成员面板或聊天记录。
        rect = session_list.rectangle()
        if rect.width() <= 0 or rect.height() <= 0:
            failed_count += 1
            print("会话列表区域没有有效屏幕范围，结束扫描。")
            break
        scroll_x = (rect.left + rect.right) // 2
        scroll_y = (rect.top + rect.bottom) // 2
        if not point_in_work_area(scroll_x, scroll_y, work_area):
            failed_count += 1
            print("会话列表滚动目标位于系统工作区之外，已停止以避开任务栏。")
            break
        pyautogui.moveTo(scroll_x, scroll_y)
        pyautogui.scroll(-6, x=scroll_x, y=scroll_y)
        time.sleep(WAIT_SECONDS)
        scroll_count += 1

    if checked_chat_count >= MAX_CHATS:
        print(f"结束原因：已达到最多检查会话数 {MAX_CHATS}（包含私聊及未识别会话）。")
    elif opened_count >= MAX_GROUPS:
        print(f"结束原因：已达到最多处理群数 {MAX_GROUPS}。")
    elif no_progress >= NO_PROGRESS_LIMIT:
        print("结束原因：连续滚动未发现新会话。")
    elif scroll_count > MAX_SCROLLS:
        print(f"结束原因：已达到会话滚动次数上限 {MAX_SCROLLS}。")
    if opened_count == 0:
        print("本次未打开任何群成员列表，因此未采集号码；请确认目标聊天可正常打开，并查看前面的群标题识别提示。")
    print_summary(
        opened_count,
        non_group_count,
        failed_count,
        excluded_chat_count,
        checked_chat_count,
    )
    print(f"点击普通成员总数：{member_clicked_count}")
    print(f"跳过群主、管理员或指定成员总数：{admin_skipped_count}")
    print(f"成员列表翻页总数：{member_page_turn_count}")
    member_number_saved_count = len(existing_numbers - initial_numbers)
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
    print(f"成功打开成员列表的群聊：{opened_count}/{MAX_GROUPS}")
    print(f"未识别到群聊标题的会话：{skipped_count}（不代表全部是私聊）")
    print(f"跳过的指定群聊：{excluded_count}")
    print(f"已检查的非排除聊天：{checked_count}/{MAX_CHATS}")
    print(f"自动处理失败数：{failed_count}")
    print("成员名称和角色仅用于识别及跳过成员；工作簿只保存去重后的数字号码。")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n收到中断信号，脚本已安全停止。")
