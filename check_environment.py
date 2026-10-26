"""Read-only checks; never click QQ controls or send requests."""

import importlib
import platform
from pathlib import Path
import sys


def main():
    root = Path(__file__).resolve().parent
    errors = []
    print(f"Python: {platform.python_version()} ({platform.architecture()[0]})")
    print(f"Interpreter: {sys.executable}")
    for name in ("pywinauto", "win32api", "win32con", "win32gui",
                 "pyautogui", "openpyxl", "odf.opendocument",
                 "rapidocr_onnxruntime", "onnxruntime", "cv2", "numpy"):
        try:
            importlib.import_module(name)
            print(f"[OK] {name}")
        except Exception as error:
            errors.append(name)
            print(f"[FAIL] {name}: {error}")
    for filename in ("zq_qq.py", "open_qq_add_friend.py", "prepare_numbers.py"):
        try:
            path = root / filename
            compile(path.read_text(encoding="utf-8-sig"), str(path), "exec")
            print(f"[OK] Syntax: {filename}")
        except Exception as error:
            errors.append(filename)
            print(f"[FAIL] {filename}: {error}")
    if errors:
        return 1
    from rapidocr_onnxruntime import RapidOCR
    try:
        import numpy
        from PIL import Image, ImageDraw, ImageFont
        engine = RapidOCR()
        sample = Image.new("RGB", (520, 100), "white")
        font_path = Path("C:/Windows/Fonts/arial.ttf")
        font = ImageFont.truetype(str(font_path), 40) if font_path.is_file() else ImageFont.load_default(size=40)
        ImageDraw.Draw(sample).text((20, 20), "QQ 123456789", font=font, fill="black")
        detections, _ = engine(numpy.asarray(sample))
        recognized = " ".join(str(item[1]) for item in detections or [])
        if "123456789" not in recognized.replace(" ", ""):
            raise RuntimeError(f"Synthetic number not recognized: {recognized}")
        print("[OK] OCR models and synthetic-number recognition")
    except Exception as error:
        errors.append("OCR")
        print(f"[FAIL] OCR: {error}")
    import open_qq_add_friend as adding
    from odf.opendocument import load
    from odf.table import Table, TableRow, TableCell
    from odf.teletype import extractText
    try:
        pending, _, _, path = adding.read_number_rows(load, Table, TableRow, TableCell, extractText)
        print(f"[OK] ODS: {path}")
        print(f"Pending numbers: {len(pending)}; required per account: {adding.NUMBERS_PER_ACCOUNT}")
        if len(pending) < adding.NUMBERS_PER_ACCOUNT:
            print("[INFO] 准备号码后再运行加好友；当前不足一个批次。")
    except Exception as error:
        errors.append("ODS")
        print(f"[FAIL] ODS: {error}")
    from pywinauto import Desktop
    import win32gui
    try:
        windows = adding.find_qq_windows(Desktop(backend="uia"), win32gui)
        print(f"QQ main windows: {len(windows)}")
        if not windows:
            print("[INFO] 请手动启动 QQ 并登录后再运行抓取或加好友。")
        for index, window in enumerate(windows, start=1):
            labels = {control.window_text().strip() for control in window.descendants()}
            print(f"Window {index}: session_list={'会话列表' in labels}; avatar={'昵称' in labels}; more={'更多' in labels}")
        print("[OK] Read-only UI Automation inspection completed")
    except Exception as error:
        print(f"[INFO] QQ 窗口检查未完成：{error}")
    print("环境检查通过。" if not errors else "环境检查存在失败项。")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
