"""Nonblocking lifecycle and logs for a disposable automation subprocess."""
from datetime import datetime
from pathlib import Path
import json
import os
import subprocess
import sys
import time
import uuid

from gui_config import state_home, app_home


class TaskController:
    def __init__(self, mode):
        self.mode = mode
        self.process = None
        self.job = None
        self.stop_event = None
        self.ready_event = None
        self.reader = None
        self.stopping_at = None
        self.forced = False
        self.log_path = None
        self.result_path = None

    @property
    def running(self):
        return self.process is not None and self.process.poll() is None

    def start(self, config, simulate=False):
        if self.running:
            raise RuntimeError('任务已经在运行。')
        self.cleanup()
        import win32event
        import win32job
        import win32api
        token = uuid.uuid4().hex
        run_dir = state_home() / 'runs' / (datetime.now().strftime('%Y%m%d_%H%M%S_') + token[:8])
        run_dir.mkdir(parents=True)
        config_file = run_dir / 'settings.json'
        config_file.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding='utf-8')
        self.log_path = run_dir / '运行日志.txt'
        self.result_path = run_dir / 'result.json'
        self.log_path.touch()
        stop_name = 'Local\\QQTaskStop_' + token
        ready_name = 'Local\\QQTaskReady_' + token
        self.stop_event = win32event.CreateEvent(None, True, False, stop_name)
        self.ready_event = win32event.CreateEvent(None, True, False, ready_name)
        self.job = win32job.CreateJobObject(None, '')
        info = win32job.QueryInformationJobObject(self.job, win32job.JobObjectExtendedLimitInformation)
        info['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(self.job, win32job.JobObjectExtendedLimitInformation, info)
        if getattr(sys, 'frozen', False):
            command = [sys.executable]
        else:
            script = app_home() / ('capture_app.py' if self.mode == 'capture' else 'add_app.py')
            command = [sys.executable, str(script)]
        command += ['--worker', str(config_file), stop_name, ready_name, str(self.log_path), str(self.result_path)]
        if simulate:
            command.append('--simulate')
        environment = os.environ.copy()
        environment['PYTHONUTF8'] = '1'
        environment['PYTHONUNBUFFERED'] = '1'
        self.stopping_at = None
        self.forced = False
        try:
            self.process = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                            stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW,
                                            env=environment, cwd=str(app_home()))
            handle = win32api.OpenProcess(0x0001 | 0x0100, False, self.process.pid)
            try:
                win32job.AssignProcessToJobObject(self.job, handle)
            finally:
                win32api.CloseHandle(handle)
            self.reader = self.log_path.open('r', encoding='utf-8', errors='replace')
            win32event.SetEvent(self.ready_event)
        except BaseException:
            if self.process and self.process.poll() is None:
                self.process.kill()
                self.process.wait(timeout=5)
            self.cleanup()
            raise

    def stop(self):
        if self.running and self.stopping_at is None:
            import win32event
            win32event.SetEvent(self.stop_event)
            self.stopping_at = time.monotonic()

    def poll(self):
        if self.running and self.stopping_at is not None and time.monotonic() - self.stopping_at >= 2.0:
            self.force_stop()
        text = self.reader.read() if self.reader else ''
        if self.process is None or self.running:
            return text, None
        code = self.process.returncode
        if self.result_path and self.result_path.is_file():
            try:
                result = json.loads(self.result_path.read_text(encoding='utf-8'))
            except (OSError, ValueError):
                result = None
        else:
            result = None
        if result is None:
            stopped = self.stopping_at is not None
            result = {'status': 'stopped' if stopped else 'error',
                      'message': '任务已强制停止，已保存的数据保留。' if stopped else f'后台任务异常退出（{code}），请查看日志。'}
        result['exit_code'] = code
        result['forced'] = self.forced
        self.cleanup()
        return text, result

    def force_stop(self):
        if self.running:
            import win32job
            self.forced = True
            if self.job:
                win32job.TerminateJobObject(self.job, 130)
            else:
                self.process.kill()

    def cleanup(self):
        import win32api
        if self.running:
            self.force_stop()
            self.process.wait(timeout=5)
        for name in ('stop_event', 'ready_event', 'job'):
            handle = getattr(self, name)
            if handle is not None:
                win32api.CloseHandle(handle)
                setattr(self, name, None)
        if self.reader:
            self.reader.close()
            self.reader = None
        self.process = None
