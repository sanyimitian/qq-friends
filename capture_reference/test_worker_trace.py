"""Verify cancellation in frozen builds without tracing automation libraries."""
import sys
import unittest
from types import ModuleType
from unittest.mock import Mock

from gui_worker import make_stop_trace
from task_control import StopRequested


class WorkerTraceTests(unittest.TestCase):
    def test_library_frames_have_no_line_trace(self):
        library = ModuleType('library')
        library.sys = sys
        exec(compile('def probe():\n    return sys._getframe().f_trace\n',
                     'library.py', 'exec'), library.__dict__)
        module = ModuleType('app')
        module.library = library
        module.__file__ = 'frozen/_internal/app.py'
        exec(compile('def main():\n    return library.probe()\n',
                     'source/app.py', 'exec'), module.__dict__)
        check = Mock()
        try:
            sys.settrace(make_stop_trace(module, check))
            observed = module.main()
        finally:
            sys.settrace(None)
        self.assertIsNone(observed)
        self.assertGreater(check.call_count, 1)

    def test_frozen_module_path_does_not_disable_cancellation(self):
        module = ModuleType('app')
        module.__file__ = 'frozen/_internal/app.py'
        exec(compile('def main():\n    return 42\n', 'source/app.py', 'exec'), module.__dict__)
        check = Mock(side_effect=StopRequested)
        try:
            sys.settrace(make_stop_trace(module, check))
            with self.assertRaises(StopRequested):
                module.main()
        finally:
            sys.settrace(None)
        check.assert_called_once()


if __name__ == '__main__':
    unittest.main()
