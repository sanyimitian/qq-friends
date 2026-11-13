"""Exercise configurable scheduling with a fake clock and worker; never operate QQ."""
import json
import os
import tempfile
import unittest
from copy import deepcopy
from unittest.mock import Mock, patch

import gui_config
from desktop_app import App, REPEAT_INTERVAL_SECONDS


class FakeController:
    def __init__(self):
        self.running = False
        self.started = []
        self.result = None
        self.log_path = 'simulated.log'
        self.failure = None

    def start(self, config, simulate=False):
        if self.failure:
            raise self.failure
        assert not self.running
        self.started.append(deepcopy(config))
        self.running = True

    def finish(self, status='finished', code=0):
        self.running = False
        self.result = {'status': status, 'message': '模拟结束', 'exit_code': code}

    def poll(self):
        result, self.result = self.result, None
        return '', result

    def stop(self):
        self.finish('stopped', 130)

    def cleanup(self):
        self.running = False


class RepeatTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='qq_repeat_test_')
        self.addCleanup(self.temp.cleanup)
        environment = patch.dict(os.environ, QQ_ASSISTANT_STATE_DIR=self.temp.name)
        environment.start()
        self.addCleanup(environment.stop)
        self.clock = patch('desktop_app.time.time', return_value=1000000).start()
        self.addCleanup(patch.stopall)
        self.config = deepcopy(dict(gui_config.defaults('add'), REPEAT_EVERY_24_HOURS=True))
        self.app = App.__new__(App)
        app = self.app
        app.mode = 'add'
        app.root = Mock()
        app.controller = FakeController()
        app.closing = app.test_mode = False
        app.repeat_config = app.next_run_at = app.wait_status = None
        app.repeat_round = 0
        app.collect = Mock(return_value=self.config)
        for name in ('set_busy', 'status', 'append_log', 'running_notice', 'progress', 'stop_button'):
            setattr(app, name, Mock())

    def begin_wait(self):
        self.app.start()
        self.clock.return_value += 300
        self.app.controller.finish()
        self.app.poll()
        return self.app.next_run_at

    def test_interval_is_exactly_24_hours_after_completion(self):
        self.assertEqual(REPEAT_INTERVAL_SECONDS, 86400)
        due = self.begin_wait()
        self.assertEqual(due, 1000300 + 86400)
        self.assertFalse(self.app.controller.running)
        self.app.running_notice.pack_forget.assert_called()
        self.app.progress.stop.assert_called()
        self.assertIn('24:00:00', self.app.status.set.call_args.args[0])
        self.app.set_busy.assert_called_once_with(True)
        self.clock.return_value = due - 0.1
        self.app.poll()
        self.assertEqual(len(self.app.controller.started), 1)
        self.clock.return_value = due
        self.app.poll()
        self.assertEqual(len(self.app.controller.started), 2)
        self.assertIsNone(self.app.next_run_at)

    def test_successive_rounds_wait_again_and_keep_original_settings(self):
        self.config['REPEAT_INTERVAL_HOURS'] = 1.5
        self.app.start()
        expected = deepcopy(self.config)
        self.config['REPEAT_INTERVAL_HOURS'] = 8
        self.config['NUMBERS_PER_ACCOUNT'] = 999
        self.config['EXCLUDED_QQ_ACCOUNTS'].append('123456789')
        for round_number in range(1, 5):
            self.clock.return_value += 37
            finished_at = self.clock.return_value
            self.app.controller.finish()
            self.app.poll()
            self.assertEqual(self.app.next_run_at, finished_at + 5400)
            self.clock.return_value = self.app.next_run_at
            self.app.poll()
            self.app.poll()
            self.assertEqual(self.app.repeat_round, round_number + 1)
        self.assertEqual(self.app.controller.started, [expected] * 5)

    def test_custom_interval_controls_due_time_countdown_and_log(self):
        for hours, countdown in [(0.01, '00:00:36'), (0.5, '00:30:00'),
                                 (12, '12:00:00'), (48, '48:00:00'), (8760, '8760:00:00')]:
            with self.subTest(hours=hours):
                self.config['REPEAT_INTERVAL_HOURS'] = hours
                due = self.begin_wait()
                self.assertEqual(due, self.clock.return_value + hours * 3600)
                self.assertIn(countdown, self.app.status.set.call_args.args[0])
                self.assertIn(f'等待 {hours:g} 小时', self.app.append_log.call_args.args[0])
                started = len(self.app.controller.started)
                self.clock.return_value = due - 0.1
                self.app.poll()
                self.assertEqual(len(self.app.controller.started), started)
                self.clock.return_value = due
                self.app.poll()
                self.app.poll()
                self.assertEqual(len(self.app.controller.started), started + 1)
                self.app.stop()
                self.app.poll()

    def test_custom_interval_applies_to_failure_and_error_paths(self):
        self.config['REPEAT_INTERVAL_HOURS'] = 0.5
        self.app.controller.failure = OSError('模拟启动失败')
        with patch('tkinter.messagebox.showerror'):
            self.app.start()
        self.assertEqual(self.app.next_run_at, self.clock.return_value + 1800)
        due = self.app.next_run_at
        self.clock.return_value = due
        self.app.poll()
        self.assertEqual(self.app.next_run_at, due + 1800)
        self.app.controller.failure = None
        for status, code in [('error', 1), ('finished', 1)]:
            self.clock.return_value = self.app.next_run_at
            self.app.poll()
            self.clock.return_value += 37
            self.app.controller.finish(status, code)
            self.app.poll()
            self.assertEqual(self.app.next_run_at, self.clock.return_value + 1800)

    def test_old_config_without_interval_still_schedules_24_hours(self):
        self.config.pop('REPEAT_INTERVAL_HOURS')
        due = self.begin_wait()
        self.assertEqual(due, self.clock.return_value + 86400)

    def test_interval_validation_and_saved_settings(self):
        self.assertEqual(gui_config.validate('add', {})['REPEAT_INTERVAL_HOURS'], 24)
        self.assertNotIn('REPEAT_INTERVAL_HOURS', gui_config.defaults('capture'))
        for invalid in ('', 'abc', 'nan', 'inf', '-inf', 0, -1, 0.001, 8760.1, 1e308):
            with self.subTest(invalid=invalid):
                with self.assertRaisesRegex(ValueError, '循环间隔'):
                    gui_config.validate('add', dict(self.config, REPEAT_INTERVAL_HOURS=invalid))
        old = dict(self.config, REPEAT_EVERY_24_HOURS=False)
        old.pop('REPEAT_INTERVAL_HOURS')
        gui_config.config_path('add').write_text(json.dumps(old), encoding='utf-8')
        loaded = gui_config.load_config('add')
        self.assertEqual(loaded['REPEAT_INTERVAL_HOURS'], 24)
        self.assertFalse(loaded['REPEAT_EVERY_24_HOURS'])
        for interval in (' 0.5 ', '12', '48'):
            gui_config.save_config('add', dict(self.config, REPEAT_INTERVAL_HOURS=interval))
            self.assertEqual(gui_config.load_config('add')['REPEAT_INTERVAL_HOURS'], float(interval))

    def test_sleep_resume_starts_only_one_round_when_overdue(self):
        due = self.begin_wait()
        self.clock.return_value = due + 5 * 86400
        for _ in range(4):
            self.app.poll()
        self.assertEqual(len(self.app.controller.started), 2)

    def test_start_is_ignored_during_run_and_wait(self):
        self.app.start()
        self.app.start()
        self.app.controller.finish()
        self.app.poll()
        due = self.app.next_run_at
        self.app.start()
        self.assertEqual(self.app.next_run_at, due)
        self.assertEqual(len(self.app.controller.started), 1)

    def test_stop_wait_cancels_future_runs_and_allows_manual_restart(self):
        due = self.begin_wait()
        self.app.stop()
        self.assertIsNone(self.app.next_run_at)
        self.app.set_busy.assert_called_with(False)
        self.clock.return_value = due + 1
        self.app.poll()
        self.assertEqual(len(self.app.controller.started), 1)
        self.app.start()
        self.assertEqual(len(self.app.controller.started), 2)
        self.assertEqual(self.app.repeat_round, 1)

    def test_stop_running_does_not_schedule(self):
        self.app.start()
        self.app.stop()
        self.app.poll()
        self.assertIsNone(self.app.next_run_at)
        self.assertIsNone(self.app.repeat_config)
        self.app.set_busy.assert_called_with(False)

    def test_stop_after_worker_exits_before_result_poll_prevents_repeat(self):
        self.app.start()
        self.app.controller.finish()
        self.app.stop()
        self.app.poll()
        self.assertIsNone(self.app.next_run_at)
        self.assertEqual(len(self.app.controller.started), 1)

    def test_close_wait_cancels_future_runs(self):
        due = self.begin_wait()
        self.app.close()
        self.clock.return_value = due + 1
        self.app.update_repeat_wait()
        self.assertIsNone(self.app.next_run_at)
        self.assertEqual(len(self.app.controller.started), 1)
        self.app.root.destroy.assert_called_once()

    def test_close_running_prevents_repeat(self):
        self.app.start()
        self.app.close()
        self.app.poll()
        self.assertIsNone(self.app.repeat_config)
        self.app.root.destroy.assert_called_once()

    def test_error_or_abnormal_exit_schedules_repeat(self):
        for status, code in [('error', 1), ('finished', 1)]:
            with self.subTest(status=status, code=code):
                self.app.start()
                self.app.controller.finish(status, code)
                self.app.poll()
                self.assertEqual(self.app.next_run_at, self.clock.return_value + 86400)
                self.clock.return_value = self.app.next_run_at
                self.app.poll()
                self.assertTrue(self.app.controller.running)

    def test_next_start_failure_waits_another_24_hours(self):
        due = self.begin_wait()
        self.app.controller.failure = OSError('模拟启动失败')
        self.clock.return_value = due
        self.app.poll()
        self.assertEqual(self.app.next_run_at, due + 86400)
        self.assertIsNotNone(self.app.repeat_config)
        self.assertIn('等待第 3 轮', self.app.status.set.call_args.args[0])
        self.app.stop()
        self.assertIsNone(self.app.next_run_at)

    def test_first_start_failure_waits_24_hours(self):
        self.app.controller.failure = OSError('模拟启动失败')
        with patch('tkinter.messagebox.showerror'):
            self.app.start()
        self.assertEqual(self.app.next_run_at, self.clock.return_value + 86400)
        self.assertIsNotNone(self.app.repeat_config)

    def test_disabled_repeat_preserves_single_run(self):
        self.config['REPEAT_EVERY_24_HOURS'] = False
        self.begin_wait()
        self.assertIsNone(self.app.next_run_at)
        self.app.set_busy.assert_called_with(False)

    def test_repeat_preserves_fill_only_mode(self):
        self.config['AUTO_SEND'] = False
        due = self.begin_wait()
        self.clock.return_value = due
        self.app.poll()
        self.assertEqual(len(self.app.controller.started), 2)
        self.assertFalse(self.app.controller.started[-1]['AUTO_SEND'])

    def test_capture_never_repeats(self):
        self.app.mode = 'capture'
        self.app.collect.return_value = gui_config.defaults('capture')
        self.begin_wait()
        self.assertIsNone(self.app.next_run_at)

    def test_old_settings_default_on_and_new_setting_round_trips(self):
        self.assertTrue(gui_config.validate('add', {})['REPEAT_EVERY_24_HOURS'])
        self.assertNotIn('REPEAT_EVERY_24_HOURS', gui_config.defaults('capture'))
        gui_config.save_config('add', self.config)
        self.assertTrue(gui_config.load_config('add')['REPEAT_EVERY_24_HOURS'])
        gui_config.save_config('add', dict(self.config, REPEAT_EVERY_24_HOURS=False))
        self.assertFalse(gui_config.load_config('add')['REPEAT_EVERY_24_HOURS'])
        with self.assertRaises(ValueError):
            gui_config.validate('add', dict(self.config, REPEAT_EVERY_24_HOURS='true'))


if __name__ == '__main__':
    unittest.main(verbosity=2)
