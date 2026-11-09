"""Exercise the real scan loop with saved settings and simulated QQ only."""
from contextlib import ExitStack, redirect_stdout
import io
import os
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import gui_config
import zq_qq_001 as capture


class ExcludedGroupTests(unittest.TestCase):
    def run_scan(self, exclusions, pages, retry_name=None):
        with ExitStack() as stack:
            folder = stack.enter_context(tempfile.TemporaryDirectory())
            stack.enter_context(patch.dict(os.environ, QQ_ASSISTANT_STATE_DIR=folder))
            config = dict(gui_config.defaults('capture'), EXCLUDED_CHAT_NAMES=exclusions)
            gui_config.save_config('capture', config)
            configured = SimpleNamespace()
            gui_config.apply_config('capture', configured, gui_config.load_config('capture'))
            stack.enter_context(patch.object(capture, 'EXCLUDED_CHAT_NAMES', configured.EXCLUDED_CHAT_NAMES))
            for name in ('SELECTED_PROCESS_ID', '_capture_workbook', '_capture_worksheet', '_capture_progress'):
                stack.enter_context(patch.object(capture, name, None))
            output = stack.enter_context(redirect_stdout(io.StringIO()))
            main = Mock()
            page_index = 0
            attempts = []
            collected = []

            def scroll(*args):
                nonlocal page_index
                page_index += 1
                return (True, '下一页') if page_index < len(pages) else (False, '已到末尾')

            def open_profile(desktop, name, *args):
                attempts.append(name)
                if name == retry_name and attempts.count(name) == 1:
                    raise RuntimeError('模拟列表临时重排')
                return main, '123456'

            def collect(*args):
                collected.append(args[-1]['group_name'])
                return (1, 0, 0, 1, 0, 0)

            mocks = {
                'load_dependencies': Mock(return_value=tuple(Mock() for _ in range(8))),
                'select_qq_window_by_number': Mock(return_value=main),
                'open_number_workbook': Mock(return_value=(Mock(), Mock(), set(), set(), [])),
                'reconcile_journal': Mock(return_value=False),
                'CaptureProgress': Mock(),
                'completed_groups_for_resume': Mock(return_value=set()),
                'set_pause_flush': Mock(), 'check_pause': Mock(),
                'find_main_qq_window': Mock(return_value=main),
                'activate_window': Mock(), 'get_work_area': Mock(return_value=(0, 0, 1200, 800)),
                'ensure_joined_groups_page': Mock(return_value=main),
                'reset_joined_groups_to_top': Mock(return_value=main),
                'get_contact_group_list': Mock(return_value=Mock()),
                'collect_joined_group_rows': Mock(side_effect=lambda *args: [(name, Mock()) for name in pages[page_index]]),
                'scroll_joined_group_list': Mock(side_effect=scroll),
                'open_group_profile_from_joined_list': Mock(side_effect=open_profile),
                'enter_group_chat_from_profile': Mock(return_value=(main, Mock())),
                'member_panel_is_open': Mock(return_value=True),
                'wait_for_member_panel': Mock(return_value=main),
                'append_group_info': Mock(side_effect=lambda *args: {}),
                'resume_for_new_task': Mock(return_value={}),
                'click_non_admin_members': Mock(side_effect=collect),
                'flush_capture': Mock(), 'sync_journal': Mock(),
                'load_journal_records': Mock(return_value=[]),
            }
            for name, mock in mocks.items():
                stack.enter_context(patch.object(capture, name, mock))
            capture.main()
            return attempts, collected, output.getvalue()

    def test_saved_multiline_exclusions_skip_before_opening_profile(self):
        attempts, collected, log = self.run_scan(
            '  AI资源共享群  \r\nCodebotAPI客服2️⃣群\n\nAI资源共享群',
            [['AI资源共享群', '正常群', 'CodebotAPI客服2️⃣群']])
        self.assertEqual(attempts, ['正常群'])
        self.assertEqual(collected, ['正常群'])
        self.assertIn('跳过的指定群聊：2', log)
        self.assertIn('按列表顺序检查的群聊：1', log)

    def test_exact_match_keeps_similar_group_names(self):
        attempts, collected, _ = self.run_scan(['资源群'], [['资源群', '资源群2', '新资源群']])
        self.assertEqual(attempts, ['资源群2', '新资源群'])
        self.assertEqual(collected, attempts)

    def test_all_excluded_pages_advance_and_overlaps_count_once(self):
        attempts, collected, log = self.run_scan(['甲群', '乙群'], [['甲群'], ['甲群', '乙群']])
        self.assertEqual(attempts, [])
        self.assertEqual(collected, [])
        self.assertIn('跳过的指定群聊：2', log)
        self.assertIn('按列表顺序检查的群聊：0', log)
        self.assertIn('自动处理失败数：0', log)

    def test_cleared_exclusions_process_every_group(self):
        attempts, collected, log = self.run_scan([], [['AI资源共享群', '普通群']])
        self.assertEqual(attempts, ['AI资源共享群', '普通群'])
        self.assertEqual(collected, attempts)
        self.assertIn('跳过的指定群聊：0', log)

    def test_other_group_can_retry_without_recounting_excluded_group(self):
        attempts, collected, log = self.run_scan(['跳过群'], [['跳过群', '重试群']], retry_name='重试群')
        self.assertEqual(attempts, ['重试群', '重试群'])
        self.assertEqual(collected, ['重试群'])
        self.assertIn('跳过的指定群聊：1', log)
        self.assertIn('按列表顺序检查的群聊：1', log)


if __name__ == '__main__':
    unittest.main()
