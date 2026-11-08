import io
import json
from contextlib import ExitStack, redirect_stdout
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
import zq_qq_001 as capture
from capture_progress import CaptureProgress, CaptureStorageError
from test_capture_navigation import member_scene, Clock, WORK_AREA


class CaptureResumeTests(unittest.TestCase):
    def test_interrupted_collection_reloads_disk_checkpoint_and_continues_next_member(self):
        from task_control import StopRequested
        main, panel, rows = member_scene(['甲', '乙', '丙'])
        clicked = []
        with tempfile.TemporaryDirectory() as directory, ExitStack() as stack:
            path = Path(directory) / 'resume.xlsx'
            stack.enter_context(patch.object(capture, 'NUMBER_WORKBOOK_PATH', path))
            stack.enter_context(patch.object(capture, 'NUMBER_JOURNAL_PATH', path.with_suffix('.jsonl')))
            stack.enter_context(patch.object(capture, 'find_main_qq_window', return_value=main))
            stack.enter_context(patch.object(capture, 'time', Clock()))
            stack.enter_context(patch.object(capture, 'MAX_MEMBERS_PER_GROUP', 3))
            stack.enter_context(patch.object(capture, 'collect_qq_snapshot', return_value={}))
            stack.enter_context(patch.object(capture, 'click_control_in_work_area',
                side_effect=lambda row, *args, **kwargs: clicked.append(row.name) or True))
            stack.enter_context(patch.object(capture, 'extract_member_number_from_popup',
                side_effect=[('12345678', '', {}), ('23456789', '', {}), ('34567890', '', {})]))
            stack.enter_context(redirect_stdout(io.StringIO()))
            workbook, summary, numbers, _, _ = capture.open_number_workbook(Workbook, load_workbook)
            section = capture.append_group_info(workbook, summary, '续采群', '777777')
            section.update(progress=CaptureProgress(path, '12345'), group_name='续采群')
            with patch.object(capture, 'check_pause', side_effect=[False, False, StopRequested()]):
                with self.assertRaises(StopRequested):
                    capture.click_non_admin_members(main, Mock(), Mock(), WORK_AREA, Mock(), Mock(),
                                                    {}, workbook, summary, numbers, section)
            progress = CaptureProgress(path, '12345')
            self.assertEqual(progress.get('777777')['next_index'], 1)
            self.assertFalse(progress.get('777777')['completed'])
            workbook.close()
            workbook, summary, numbers, completed, records = capture.open_number_workbook(Workbook, load_workbook)
            capture.reconcile_journal(workbook, summary, records)
            self.assertNotIn('777777', completed)
            section = capture.append_group_info(workbook, summary, '续采群', '777777')
            section.update(progress=progress, group_name='续采群', resume=progress.get('777777'),
                           saved_count=capture.group_saved_count(summary, '777777'))
            capture.click_non_admin_members(main, Mock(), Mock(), WORK_AREA, Mock(), Mock(),
                                            {}, workbook, summary, numbers, section)
            self.assertEqual(clicked, ['甲', '乙', '丙'])
            self.assertTrue(section['completed'])
            self.assertEqual(list(workbook['续采群_777777'].values)[1:],
                             [('12345678', None), ('23456789', None), ('34567890', None)])
            self.assertEqual(workbook.sheetnames.count('续采群_777777'), 1)
            workbook.close()

    def test_progress_is_persistent_account_scoped_and_separate_per_file(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'numbers.xlsx'
            progress = CaptureProgress(path, '12345')
            progress.save('777777', group_name='测试群', page_names=['甲', '乙'], next_index=1,
                          scroll_percent=25, page_turns=2, saved_count=1, completed=False)
            self.assertEqual(CaptureProgress(path, '12345').get('777777')['next_index'], 1)
            self.assertEqual(CaptureProgress(path, '54321').get('777777'), {})
            self.assertEqual(CaptureProgress(path.with_name('other.xlsx'), '12345').groups, {})
            with patch('capture_progress.os.replace', side_effect=PermissionError('locked')):
                with self.assertRaises(CaptureStorageError):
                    progress.save('777777', completed=True)
            self.assertFalse(CaptureProgress(path, '12345').get('777777')['completed'])
            self.assertEqual(list(Path(directory).glob('.qq_progress_*')), [])

    def test_partial_group_with_members_is_not_completed(self):
        with tempfile.TemporaryDirectory() as directory, ExitStack() as stack:
            path = Path(directory) / 'numbers.xlsx'
            stack.enter_context(patch.object(capture, 'NUMBER_WORKBOOK_PATH', path))
            stack.enter_context(patch.object(capture, 'NUMBER_JOURNAL_PATH', path.with_suffix('.jsonl')))
            workbook, summary, numbers, _, _ = capture.open_number_workbook(Workbook, load_workbook)
            group = capture.append_group_info(workbook, summary, '部分完成', '777777')
            capture.append_number(workbook, summary, numbers, '12345678', group)
            capture.sync_journal()
            recovered, summary, numbers, completed, records = capture.open_number_workbook(Workbook, load_workbook)
            capture.reconcile_journal(recovered, summary, records)
            self.assertEqual(completed, set())
            self.assertEqual(numbers, {'12345678'})
            self.assertEqual(recovered['部分完成_777777']['A2'].value, '12345678')
            capture.append_journal_record({'type': 'group_complete', 'group_number': '777777'})
            self.assertEqual(capture.open_number_workbook(Workbook, load_workbook)[3], set())

    def test_recovery_preserves_later_group_and_processed_status(self):
        workbook = Workbook()
        summary = workbook.active
        summary.title = 'QQ号码'
        summary.append(['甲', '111111'])
        summary.append(['11111111', None, '已处理'])
        summary.append(['乙', '222222'])
        summary.append(['22222222', None, '已处理'])
        records = [
            {'type': 'group', 'group_name': '甲', 'group_number': '111111'},
            {'type': 'member', 'group_number': '111111', 'number': '33333333'},
            {'type': 'group', 'group_name': '乙', 'group_number': '222222'},
            {'type': 'member', 'group_number': '222222', 'number': '11111111'},
        ]
        with patch.object(capture, 'atomic_save'):
            capture.reconcile_journal(workbook, summary, records)
            snapshot = list(summary.values)
            self.assertFalse(capture.reconcile_journal(workbook, summary, records))
        self.assertEqual(list(summary.values), snapshot)
        self.assertEqual(summary['A3'].value, '乙')
        self.assertEqual(summary['C2'].value, '已处理')
        self.assertEqual(summary['C4'].value, '已处理')
        self.assertEqual(workbook['甲_111111']['A3'].value, '33333333')
        self.assertEqual(workbook['甲_111111']['C2'].value, '已处理')
        self.assertEqual(workbook['乙_222222']['C2'].value, '已处理')
        self.assertEqual(workbook['乙_222222']['A3'].value, '11111111')

    def test_long_invalid_and_colliding_group_names_are_valid_and_reused(self):
        workbook = Workbook()
        summary = workbook.active
        name = "'测试/*?:[]\\😀" * 10
        title = capture.group_sheet_title(name, '123456789012')
        self.assertLessEqual(len(title.encode('utf-16-le')) // 2, 31)
        self.assertTrue(title.endswith('_123456789012'))
        workbook.create_sheet(title)
        sheet, created = capture.ensure_group_sheet(workbook, summary, name, '123456789012')
        self.assertTrue(created)
        self.assertNotEqual(sheet.title, title)
        reused, created = capture.ensure_group_sheet(workbook, summary, '改名群', '123456789012')
        self.assertIs(reused, sheet)
        self.assertFalse(created)

    def test_member_resume_skips_saved_row_and_preserves_total_limit(self):
        main, panel, rows = member_scene(['甲', '乙', '丙'])
        progress = Mock()
        resume = dict(page_names=['甲', '乙', '丙'], next_index=1, page_turns=0,
                      saved_count=1, scroll_percent=None)
        section = dict(group_number='777777', group_name='测试群', progress=progress, resume=resume)
        clicked = []
        with ExitStack() as stack:
            stack.enter_context(redirect_stdout(io.StringIO()))
            stack.enter_context(patch.object(capture, 'find_main_qq_window', return_value=main))
            stack.enter_context(patch.object(capture, 'time', Clock()))
            stack.enter_context(patch.object(capture, 'MAX_MEMBERS_PER_GROUP', 2))
            stack.enter_context(patch.object(capture, 'collect_qq_snapshot', return_value={}))
            stack.enter_context(patch.object(capture, 'click_control_in_work_area',
                                            side_effect=lambda row, *args, **kwargs: clicked.append(row.name) or True))
            stack.enter_context(patch.object(capture, 'extract_member_number_from_popup', return_value=('12345678', '', {})))
            stack.enter_context(patch.object(capture, 'append_number', return_value=True))
            stack.enter_context(patch.object(capture, 'sync_journal'))
            capture.click_non_admin_members(main, Mock(), Mock(), WORK_AREA, Mock(), Mock(),
                                            {}, Mock(), Mock(), set(), section)
        self.assertEqual(clicked, ['乙'])
        self.assertTrue(section['completed'])
        self.assertEqual(progress.save.call_args.kwargs['next_index'], 2)
        self.assertEqual(progress.save.call_args.kwargs['saved_count'], 2)

    def test_changed_member_page_restarts_without_trusting_old_row_offset(self):
        main, panel, _ = member_scene(['新成员', '乙'])
        panel.iface_scroll = Mock()
        resume = dict(page_names=['甲', '乙'], next_index=1, page_turns=5, scroll_percent=50)
        with patch.object(capture, 'time', Clock()), redirect_stdout(io.StringIO()):
            self.assertEqual(capture.restore_capture_position(main, Mock(), Mock(), WORK_AREA, resume), (0, 0))
        panel.iface_scroll.SetScrollPercent.assert_called_with(-1.0, 0.0)

    def test_flush_can_recover_journal_after_interruption_before_excel_write(self):
        with tempfile.TemporaryDirectory() as directory, ExitStack() as stack:
            path = Path(directory) / 'numbers.xlsx'
            workbook = Workbook()
            summary = workbook.active
            stack.enter_context(patch.object(capture, 'NUMBER_WORKBOOK_PATH', path))
            stack.enter_context(patch.object(capture, 'NUMBER_JOURNAL_PATH', path.with_suffix('.jsonl')))
            stack.enter_context(patch.object(capture, '_capture_workbook', workbook))
            stack.enter_context(patch.object(capture, '_capture_worksheet', summary))
            capture.append_journal_record({'type': 'group', 'group_number': '777777', 'group_name': '恢复群'})
            capture.append_journal_record({'type': 'member', 'group_number': '777777', 'number': '12345678'})
            capture.flush_capture()
            check = load_workbook(path)
            self.assertEqual(check['恢复群_777777']['A2'].value, '12345678')
            check.close()


if __name__ == '__main__':
    unittest.main()
