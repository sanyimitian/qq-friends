"""Own Ctrl+B on a dedicated Windows message thread while capture runs."""
import ctypes
from ctypes import wintypes
import threading


class CaptureHotkey:
    def __init__(self, toggle):
        self.toggle = toggle
        self.ready = threading.Event()
        self.thread_id = None
        self.registered = False
        self.thread = threading.Thread(target=self._run, daemon=True)

    def start(self):
        self.thread.start()
        self.ready.wait(3)
        return self.registered

    def _run(self):
        user32 = ctypes.windll.user32
        self.thread_id = ctypes.windll.kernel32.GetCurrentThreadId()
        message = wintypes.MSG()
        user32.PeekMessageW(ctypes.byref(message), None, 0, 0, 0)
        # MOD_CONTROL | MOD_NOREPEAT prevents a held key toggling repeatedly.
        self.registered = bool(user32.RegisterHotKey(None, 1, 0x0002 | 0x4000, ord('B')))
        self.ready.set()
        if not self.registered:
            return
        try:
            while user32.GetMessageW(ctypes.byref(message), None, 0, 0) > 0:
                if message.message == 0x0312 and message.wParam == 1:
                    self.toggle()
        finally:
            user32.UnregisterHotKey(None, 1)
            self.registered = False

    def close(self):
        if self.thread_id and self.thread.is_alive():
            ctypes.windll.user32.PostThreadMessageW(self.thread_id, 0x0012, 0, 0)
            self.thread.join(timeout=2)
