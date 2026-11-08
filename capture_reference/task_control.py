"""Cancellation and atomic writes shared by desktop workers and legacy scripts."""
from pathlib import Path
import os
import tempfile
import time


class StopRequested(BaseException):
    pass


_stop_event = None
_pause_event = None
_paused_event = None
_pause_flush = None


def set_pause_events(request=None, acknowledged=None):
    global _pause_event, _paused_event, _pause_flush
    _pause_event, _paused_event = request, acknowledged
    _pause_flush = None


def set_pause_flush(callback):
    global _pause_flush
    _pause_flush = callback


def check_pause():
    """Pause only at capture transaction boundaries, after saving progress."""
    check_stop()
    if _pause_event is None:
        return False
    import win32event
    if win32event.WaitForSingleObject(_pause_event, 0) != win32event.WAIT_OBJECT_0:
        return False
    if _pause_flush is not None:
        _pause_flush()
    win32event.SetEvent(_paused_event)
    print('采集已暂停，进度已保存。按 Ctrl+B 或点击继续恢复；请保留 QQ 当前页面。', flush=True)
    try:
        while win32event.WaitForSingleObject(_pause_event, 0) == win32event.WAIT_OBJECT_0:
            check_stop()
            time.sleep(0.05)
        check_stop()
    finally:
        win32event.ResetEvent(_paused_event)
    print('继续采集。', flush=True)
    return True


def set_stop_event(event):
    global _stop_event
    _stop_event = event


def check_stop():
    if _stop_event is not None:
        import win32event
        if win32event.WaitForSingleObject(_stop_event, 0) == win32event.WAIT_OBJECT_0:
            raise StopRequested()


class InterruptibleTime:
    def __getattr__(self, name):
        return getattr(time, name)

    def sleep(self, seconds):
        if _stop_event is None:
            time.sleep(seconds)
            return
        import win32event
        if win32event.WaitForSingleObject(_stop_event, max(0, int(seconds * 1000))) == win32event.WAIT_OBJECT_0:
            raise StopRequested()


def atomic_save(document, destination, ods=False):
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, name = tempfile.mkstemp(prefix='.qq_save_', suffix=destination.suffix, dir=destination.parent)
    os.close(descriptor)
    try:
        if ods:
            document.save(name, addsuffix=False)
        else:
            document.save(name)
        os.replace(name, destination)
    finally:
        if os.path.exists(name):
            os.unlink(name)
