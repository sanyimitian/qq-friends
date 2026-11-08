"""Durable per-workbook capture positions; never infer completion from members."""
import json
import os
import tempfile
from pathlib import Path


class CaptureStorageError(BaseException):
    """Stop automation when progress cannot be committed."""


class CaptureProgress:
    def __init__(self, workbook_path, account):
        self.path = Path(workbook_path).with_suffix('.capture-progress.json')
        self.account = str(account)
        self.data = {'version': 1, 'accounts': {}}
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text(encoding='utf-8'))
                if self.data.get('version') != 1 or not isinstance(self.data.get('accounts'), dict):
                    raise ValueError('进度文件格式不正确')
                for groups in self.data['accounts'].values():
                    if not isinstance(groups, dict):
                        raise ValueError('账号进度格式不正确')
                    for position in groups.values():
                        if not isinstance(position, dict) or not isinstance(position.get('completed'), bool):
                            raise ValueError('群进度格式不正确')
                        if position.get('completed'):
                            continue
                        for key in ('next_index', 'page_turns', 'saved_count'):
                            value = position.get(key, 0)
                            if type(value) is not int or value < 0:
                                raise ValueError('成员位置格式不正确')
                        names = position.get('page_names', [])
                        if not isinstance(names, list) or any(not isinstance(name, str) for name in names):
                            raise ValueError('成员页面格式不正确')
                        if names and position.get('next_index', 0) > len(names):
                            raise ValueError('成员行超出页面范围')
                        percent = position.get('scroll_percent')
                        if percent is not None and (type(percent) not in (int, float) or not 0 <= percent <= 100):
                            raise ValueError('滚动位置格式不正确')
            except (OSError, ValueError, AttributeError) as error:
                raise CaptureStorageError(f'无法读取采集进度 {self.path}：{error}') from error
        self.groups = self.data['accounts'].setdefault(self.account, {})

    def get(self, group_number):
        return self.groups.get(str(group_number), {})

    def save(self, group_number, **position):
        self.groups[str(group_number)] = position
        self.path.parent.mkdir(parents=True, exist_ok=True)
        name = None
        try:
            fd, name = tempfile.mkstemp(prefix='.qq_progress_', suffix='.json', dir=self.path.parent)
            with os.fdopen(fd, 'w', encoding='utf-8') as stream:
                json.dump(self.data, stream, ensure_ascii=False, indent=2)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(name, self.path)
        except OSError as error:
            raise CaptureStorageError(f'采集进度保存失败，已停止操作：{error}') from error
        finally:
            if name and os.path.exists(name):
                os.unlink(name)
