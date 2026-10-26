"""Preflight and durable recovery for friend requests; never interact with QQ."""
from pathlib import Path
import hashlib
import json
import os
import re
import stat
import sys
import tempfile

from gui_config import state_home


def ensure_number_file_writable(path):
    """Check replacement rights without writing to or renaming the workbook."""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f'找不到号码文件：{path}')
    temporary = None
    renamed = None
    try:
        if not path.stat().st_mode & stat.S_IWRITE:
            raise PermissionError('号码文件是只读文件。')
        if sys.platform == 'win32':
            import win32con
            import win32file

            handle = win32file.CreateFile(
                str(path),
                win32con.GENERIC_READ | win32con.GENERIC_WRITE | 0x00010000,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                None, win32con.OPEN_EXISTING, 0, None,
            )
            handle.Close()
        else:
            # The application runs on Windows; this keeps offline checks useful elsewhere.
            with path.open('r+b'):
                pass
        descriptor, name = tempfile.mkstemp(prefix='.qq_write_check_', dir=path.parent)
        temporary = Path(name)
        os.close(descriptor)
        renamed = temporary.with_name(temporary.name + '.rename')
        os.replace(temporary, renamed)
        renamed.unlink()
        renamed = None
    except Exception as error:
        raise PermissionError(
            f'号码表暂时无法写入，请关闭 Excel/WPS 中的该文件，取消文件只读属性，'
            f'并确认文件夹可写后再开始：{path}。原因：{error}'
        ) from error
    finally:
        for candidate in (temporary, renamed):
            if candidate is not None:
                try:
                    candidate.unlink(missing_ok=True)
                except OSError:
                    pass


def _canonical_source(path):
    return os.path.normcase(str(Path(path).resolve()))


class AttemptJournal:
    """Remember uncertain sends and completed attempts until the workbook is saved.

    ``all_numbers`` supplied to recover must contain every valid QQ number in the
    current sheet, including rows whose B status is already marked.  Row positions
    in the journal are diagnostic only and are never used to recover workbook data.
    """

    def __init__(self, number_file, sheet_name, processed_marker='已处理'):
        self.source = _canonical_source(number_file)
        self.sheet = str(sheet_name)
        self.processed_marker = processed_marker
        directory = state_home() / 'add_recovery'
        directory.mkdir(parents=True, exist_ok=True)
        digest = hashlib.sha256(self.source.encode('utf-8')).hexdigest()
        self.path = directory / f'{digest}.json'
        self.records = {}
        if self.path.exists():
            self._load()

    def _invalid(self, reason):
        return RuntimeError(f'好友处理恢复记录无法安全读取，已停止，请先核对恢复文件：{self.path}。{reason}')

    def _load(self):
        try:
            payload = json.loads(self.path.read_text(encoding='utf-8'))
        except (OSError, ValueError) as error:
            raise self._invalid(str(error)) from error
        if not isinstance(payload, dict) or payload.get('version') != 1:
            raise self._invalid('记录格式或版本不正确。')
        if payload.get('source') != self.source:
            raise self._invalid('记录对应的号码文件不一致。')
        records = payload.get('records')
        if not isinstance(records, dict) or not isinstance(payload.get('sheet'), str):
            raise self._invalid('记录缺少工作表或号码信息。')
        for number, record in records.items():
            if not isinstance(record, dict):
                raise self._invalid('号码记录格式不正确。')
            try:
                self._validate_record(number, record.get('rows'), record.get('state'))
            except ValueError as error:
                raise self._invalid(str(error)) from error
        if records and payload['sheet'] != self.sheet:
            raise self._invalid(
                f'仍有未恢复记录属于工作表“{payload["sheet"]}”，当前选择的是“{self.sheet}”；'
                '请切回原工作表处理。'
            )
        self.records = records

    @staticmethod
    def _validate_record(number, rows, state):
        if not isinstance(number, str) or re.fullmatch(r'[0-9]{5,12}', number) is None:
            raise ValueError('恢复记录中的 QQ 号不正确。')
        if (not isinstance(rows, list) or not rows
                or any(type(row) is not int or row < 1 for row in rows)):
            raise ValueError('恢复记录中的行号不正确。')
        if state not in ('sending', 'processed'):
            raise ValueError('恢复记录中的处理阶段不正确。')

    def _persist(self, records):
        payload = {'version': 1, 'source': self.source, 'sheet': self.sheet, 'records': records}
        descriptor, name = tempfile.mkstemp(prefix='.qq_recovery_', suffix='.tmp', dir=self.path.parent)
        try:
            with os.fdopen(descriptor, 'w', encoding='utf-8', newline='\n') as stream:
                json.dump(payload, stream, ensure_ascii=False, indent=2)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(name, self.path)
        except Exception as error:
            raise RuntimeError(f'无法保存好友处理恢复记录，已停止：{self.path}。原因：{error}') from error
        finally:
            try:
                Path(name).unlink(missing_ok=True)
            except OSError:
                pass

    def record(self, number, rows, state):
        rows = list(rows)
        self._validate_record(number, rows, state)
        if self.records.get(number, {}).get('state') == 'processed' and state == 'sending':
            raise RuntimeError(f'QQ {number} 已有待写回的处理记录，请先恢复号码表。')
        records = dict(self.records)
        records[number] = {'rows': rows, 'state': state}
        self._persist(records)
        self.records = records

    def clear(self, number):
        if number not in self.records:
            return
        records = dict(self.records)
        del records[number]
        if records:
            self._persist(records)
        else:
            try:
                self.path.unlink(missing_ok=True)
            except OSError as error:
                raise RuntimeError(f'号码表已保存，但恢复记录暂时无法清理：{self.path}。请再次开始以完成核对。') from error
        self.records = records

    def recover(self, numbers, mark_callback, all_numbers=None):
        """Recover known attempts; block uncertain sends before any QQ interaction."""
        pending = dict(numbers)
        known = set(pending) if all_numbers is None else set(all_numbers)
        if not set(pending).issubset(known):
            raise ValueError('all_numbers 必须包含当前工作表所有待处理和已标记的 QQ 号。')
        # Validate every record first so an uncertain/missing number does not cause
        # unrelated workbook edits before the user sees the error.
        for number, record in self.records.items():
            if number not in known:
                raise RuntimeError(
                    f'恢复记录中的 QQ {number} 已不在当前工作表，已停止且未按旧行号写入。'
                    f'请核对号码表和恢复文件：{self.path}'
                )
            if number in pending and record['state'] == 'sending':
                raise RuntimeError(
                    f'QQ {number} 上次可能已点击发送，程序不会自动重复发送。请先在 QQ 中核对；'
                    f'如已发送或希望跳过，请将该号码 B 列标为“{self.processed_marker}”并保存关闭表格；'
                    f'仅在确认未发送且需要重试时，核对并删除本地恢复文件：{self.path}，然后重新开始。'
                )
        recovered = set()
        for number in list(self.records):
            if number in pending:
                mark_callback(pending[number])
                recovered.add(number)
            self.clear(number)
        return [(number, rows) for number, rows in numbers if number not in recovered]
