"""Offline packaged-dependency check. Does not inspect or operate QQ."""
from pathlib import Path
import json
import tempfile


def run(output):
    report = {}
    try:
        import pyautogui
        import win32api
        import win32gui
        import win32job
        import win32event
        from pywinauto import Desktop
        import zq_qq
        import open_qq_add_friend
        import numpy
        import cv2
        from PIL import Image, ImageDraw, ImageFont
        from rapidocr_onnxruntime import RapidOCR
        from openpyxl import Workbook
        from prepare_numbers import sync_numbers
        from odf.opendocument import load
        from odf.table import Table, TableRow, TableCell
        from odf.teletype import extractText
        from odf.text import P
        report['imports'] = True
        engine = RapidOCR()
        picture = Image.new('RGB', (520, 100), 'white')
        font = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 40)
        ImageDraw.Draw(picture).text((20, 20), 'QQ 123456789', fill='black', font=font)
        detections, _ = engine(numpy.asarray(picture))
        text = ' '.join(str(item[1]) for item in detections or [])
        if '123456789' not in text.replace(' ', ''):
            raise RuntimeError('OCR test failed: ' + text)
        report['ocr'] = True
        with tempfile.TemporaryDirectory(prefix='qq_package_check_') as folder:
            source, destination = Path(folder) / 'input.xlsx', Path(folder) / 'numbers.ods'
            workbook = Workbook()
            workbook.active.append(['123456789'])
            workbook.save(source)
            workbook.close()
            converted = sync_numbers(source, destination)
            assert converted['added'] == 1
            original_path = open_qq_add_friend.NUMBER_FILE_PATH
            try:
                open_qq_add_friend.NUMBER_FILE_PATH = destination
                rows, *_ = open_qq_add_friend.read_number_rows(load, Table, TableRow, TableCell, extractText)
                assert rows == [('123456789', [1])]
                open_qq_add_friend.NUMBER_FILE_PATH = source
                rows, document, sheet, path = open_qq_add_friend.read_number_rows(load, Table, TableRow, TableCell, extractText)
                try:
                    assert rows == [('123456789', [1])]
                    open_qq_add_friend.mark_rows_processed(document, sheet, [1], TableRow, TableCell, P, path)
                finally:
                    document.close()
                rows, document, sheet, _ = open_qq_add_friend.read_number_rows(load, Table, TableRow, TableCell, extractText)
                try:
                    assert rows == []
                    assert sheet['B1'].value == open_qq_add_friend.PROCESSED_MARKER
                finally:
                    document.close()
                report['xlsx_direct_read_write'] = True
            finally:
                open_qq_add_friend.NUMBER_FILE_PATH = original_path
        report['xlsx_ods'] = True
        report['ok'] = True
    except BaseException as error:
        import traceback
        report['ok'] = False
        report['error'] = str(error)
        report['traceback'] = traceback.format_exc()
    Path(output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0 if report.get('ok') else 1
