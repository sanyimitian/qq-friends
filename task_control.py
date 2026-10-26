"""Cancellation and atomic writes shared by desktop workers and legacy scripts."""
from pathlib import Path
import os
import tempfile
import time


class StopRequested(BaseException):
    pass


_stop_event = None


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
