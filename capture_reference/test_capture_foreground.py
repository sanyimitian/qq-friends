"""Activation regressions: simulated windows only, no QQ interaction."""
import io
from contextlib import redirect_stdout
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
import zq_qq_001 as capture
from test_capture_detection import Clock
from task_control import StopRequested

class ForegroundTests(unittest.TestCase):
    def scene(self):
        window = Mock(handle=123)
        gui = Mock()
        gui.IsWindow.return_value = True
        gui.IsIconic.return_value = False
        gui.IsWindowVisible.return_value = True
        gui.GetForegroundWindow.return_value = 999
        gui.GetWindowPlacement.return_value = (0, 3, None, None, None)
        con = SimpleNamespace(SW_RESTORE=9, SW_MAXIMIZE=3)
        return window, gui, con

    def test_setforeground_error_falls_back_to_uia_without_fixed_wait(self):
        window, gui, con = self.scene()
        gui.SetForegroundWindow.side_effect = OSError(2, 'SetForegroundWindow')
        window.set_focus.side_effect = lambda: setattr(gui.GetForegroundWindow, 'return_value', 123)
        clock = Clock()
        with patch.object(capture, 'time', clock):
            capture.activate_window(window, con, gui)
        window.set_focus.assert_called_once()
        self.assertEqual(clock.sleeps, [])

    def test_exception_after_focus_was_set_does_not_fail(self):
        window, gui, con = self.scene()
        gui.GetForegroundWindow.return_value = 123
        gui.SetForegroundWindow.side_effect = OSError(2, 'SetForegroundWindow')
        with patch.object(capture, 'time', Clock()):
            capture.activate_window(window, con, gui)
        window.set_focus.assert_not_called()

    def test_silent_foreground_failure_also_uses_focus(self):
        window, gui, con = self.scene()
        window.set_focus.side_effect = lambda: setattr(gui.GetForegroundWindow, 'return_value', 123)
        with patch.object(capture, 'time', Clock()):
            capture.activate_window(window, con, gui)
        window.set_focus.assert_called_once()

    def test_wrong_foreground_times_out_with_original_budget(self):
        window, gui, con = self.scene()
        clock = Clock()
        with patch.object(capture, 'time', clock), self.assertRaisesRegex(RuntimeError, '无法激活所选 QQ'):
            capture.activate_window(window, con, gui)
        self.assertAlmostEqual(sum(clock.sleeps), .8)

    def test_invalid_handle_never_calls_focus_or_click(self):
        window, gui, con = self.scene()
        gui.IsWindow.return_value = False
        with self.assertRaisesRegex(RuntimeError, '主窗口已失效'):
            capture.activate_window(window, con, gui)
        window.set_focus.assert_not_called()
        gui.SetForegroundWindow.assert_not_called()

    def test_maximize_uses_same_fallback_and_existing_timeout(self):
        window, gui, con = self.scene()
        gui.SetForegroundWindow.side_effect = OSError(2, 'SetForegroundWindow')
        window.set_focus.side_effect = lambda: setattr(gui.GetForegroundWindow, 'return_value', 123)
        clock = Clock()
        with patch.object(capture, 'time', clock):
            capture.activate_window(window, con, gui, timeout=1, maximize=True)
        gui.ShowWindow.assert_called_with(123, con.SW_MAXIMIZE)
        self.assertEqual(clock.sleeps, [])

    def test_stop_is_not_swallowed(self):
        window, gui, con = self.scene()
        gui.SetForegroundWindow.side_effect = StopRequested()
        with self.assertRaises(StopRequested):
            capture.activate_window(window, con, gui)

    def test_startup_account_error_is_propagated_before_workbook_open(self):
        dependencies = tuple(Mock() for _ in range(8))
        with patch.object(capture, 'load_dependencies', return_value=dependencies), patch.object(
                capture, 'select_qq_window_by_number', side_effect=RuntimeError('activation failed')), patch.object(
                capture, 'open_number_workbook') as opened, redirect_stdout(io.StringIO()):
            with self.assertRaisesRegex(RuntimeError, 'activation failed'):
                capture.main()
        opened.assert_not_called()

if __name__ == '__main__':
    unittest.main()

