"""Offline checks for editable friend-request verification settings."""
import json
import os
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import gui_config


EXPECTED_DEFAULT = '您好，0.18稳定不降智中转免费送额度，欢迎来试，地址看头像'


class VerificationSettingsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='qq_verification_settings_')
        self.addCleanup(self.temp.cleanup)
        self.environment = patch.dict(os.environ, {'QQ_ASSISTANT_STATE_DIR': self.temp.name})
        self.environment.start()
        self.addCleanup(self.environment.stop)

    def test_default_field_is_add_only_and_near_auto_send(self):
        fields = gui_config.fields('add')
        keys = [field.key for field in fields]
        field = fields[keys.index('VERIFICATION_MESSAGE')]
        self.assertEqual(field.default, EXPECTED_DEFAULT)
        self.assertEqual(field.kind, 'multiline')
        self.assertEqual(field.label, '加好友验证信息')
        self.assertEqual(field.group, '常用设置')
        self.assertEqual(keys.index('VERIFICATION_MESSAGE') + 1, keys.index('AUTO_SEND'))
        self.assertNotIn('VERIFICATION_MESSAGE', gui_config.defaults('capture'))

    def test_old_settings_receive_new_default_without_losing_values(self):
        old = gui_config.defaults('add')
        old.pop('VERIFICATION_MESSAGE')
        old['NUMBERS_PER_ACCOUNT'] = 5
        gui_config.config_path('add').write_text(json.dumps(old), encoding='utf-8')
        loaded = gui_config.load_config('add')
        self.assertEqual(loaded['VERIFICATION_MESSAGE'], EXPECTED_DEFAULT)
        self.assertEqual(loaded['NUMBERS_PER_ACCOUNT'], 5)

    def test_custom_message_survives_save_load_and_apply(self):
        message = '您好，测试验证信息。\n第二行保持不变。'
        config = gui_config.defaults('add')
        config['VERIFICATION_MESSAGE'] = message
        gui_config.save_config('add', config)
        loaded = gui_config.load_config('add')
        self.assertEqual(loaded['VERIFICATION_MESSAGE'], message)
        module = SimpleNamespace()
        gui_config.apply_config('add', module, loaded)
        self.assertEqual(module.VERIFICATION_MESSAGE, message)
        self.assertIsInstance(module.VERIFICATION_MESSAGE, str)

    def test_empty_message_is_rejected(self):
        with self.assertRaisesRegex(ValueError, '加好友验证信息.*不能为空'):
            gui_config.validate('add', {'VERIFICATION_MESSAGE': ' \n '})

    def make_app(self):
        import tkinter as tk
        from desktop_app import App
        try:
            root = tk.Tk()
        except tk.TclError as error:
            self.skipTest(f'Tk is unavailable: {error}')
        root.withdraw()
        self.addCleanup(root.destroy)
        app = App(root, 'add')
        self.addCleanup(app.controller.cleanup)
        return app

    def test_page_loads_and_collects_custom_multiline_message(self):
        import tkinter as tk
        configured = '已保存验证信息\n第二行'
        gui_config.save_config('add', {'VERIFICATION_MESSAGE': configured})
        app = self.make_app()
        widget = app.inputs['VERIFICATION_MESSAGE']
        self.assertIsInstance(widget, tk.Text)
        self.assertEqual(widget.get('1.0', 'end-1c'), configured)
        widget.delete('1.0', 'end')
        widget.insert('1.0', '编辑后的验证信息\n保留换行')
        self.assertEqual(app.collect()['VERIFICATION_MESSAGE'], '编辑后的验证信息\n保留换行')
        self.assertFalse(app.controller.running)

    def test_reset_restores_message_and_list_fields(self):
        app = self.make_app()
        widget = app.inputs['VERIFICATION_MESSAGE']
        widget.delete('1.0', 'end')
        widget.insert('1.0', '自定义文字')
        app.inputs['RESULT_ACTIONS'].delete('1.0', 'end')
        app.inputs['RESULT_ACTIONS'].insert('1.0', '测试按钮')
        app.reset()
        collected = app.collect()
        self.assertEqual(collected['VERIFICATION_MESSAGE'], EXPECTED_DEFAULT)
        self.assertEqual(collected['RESULT_ACTIONS'], gui_config.defaults('add')['RESULT_ACTIONS'])
        self.assertFalse(app.controller.running)


if __name__ == '__main__':
    unittest.main(verbosity=2)
