"""Offline file-lock and recovery checks; use temporary files only."""
from pathlib import Path
import json
import os
import stat
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

import add_progress


class ProgressTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix='qq_progress_test_')
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        self.path = self.directory / 'numbers.xlsx'
        self.path.write_bytes(b'synthetic workbook')
        self.state = self.directory / 'state'
        self.state.mkdir()
        self.environment = patch.dict(os.environ, {'QQ_ASSISTANT_STATE_DIR': str(self.state)})
        self.environment.start()
        self.addCleanup(self.environment.stop)

    def journal(self, sheet='Sheet'):
        return add_progress.AttemptJournal(self.path, sheet)

    def test_preflight_does_not_change_original_or_leave_temp_files(self):
        before = self.path.read_bytes(), self.path.stat().st_mtime_ns
        add_progress.ensure_number_file_writable(self.path)
        self.assertEqual(before, (self.path.read_bytes(), self.path.stat().st_mtime_ns))
        self.assertEqual(list(self.directory.glob('.qq_write_check_*')), [])

    def test_preflight_rejects_missing_file(self):
        with self.assertRaisesRegex(FileNotFoundError, '找不到号码文件'):
            add_progress.ensure_number_file_writable(self.directory / 'missing.xlsx')

    def test_preflight_rejects_readonly_file(self):
        self.path.chmod(stat.S_IREAD)
        try:
            with self.assertRaisesRegex(PermissionError, '只读'):
                add_progress.ensure_number_file_writable(self.path)
        finally:
            self.path.chmod(stat.S_IREAD | stat.S_IWRITE)

    def test_preflight_reports_directory_create_failure(self):
        before = self.path.read_bytes(), self.path.stat().st_mtime_ns
        with patch.object(add_progress.tempfile, 'mkstemp', side_effect=PermissionError('folder denied')):
            with self.assertRaisesRegex(PermissionError, 'Excel/WPS.*folder denied'):
                add_progress.ensure_number_file_writable(self.path)
        self.assertEqual(before, (self.path.read_bytes(), self.path.stat().st_mtime_ns))

    def test_preflight_cleans_temporary_after_rename_failure(self):
        with patch.object(add_progress.os, 'replace', side_effect=PermissionError('rename denied')):
            with self.assertRaisesRegex(PermissionError, 'rename denied'):
                add_progress.ensure_number_file_writable(self.path)
        self.assertEqual(list(self.directory.glob('.qq_write_check_*')), [])
        self.assertEqual(self.path.read_bytes(), b'synthetic workbook')

    @unittest.skipUnless(sys.platform == 'win32', 'Windows sharing semantics')
    def test_preflight_detects_actual_windows_replacement_lock(self):
        import win32con
        import win32file
        before = self.path.read_bytes(), self.path.stat().st_mtime_ns
        blocker = win32file.CreateFile(
            str(self.path), win32con.GENERIC_READ | win32con.GENERIC_WRITE,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
            None, win32con.OPEN_EXISTING, 0, None,
        )
        try:
            with self.assertRaisesRegex(PermissionError, '关闭 Excel/WPS'):
                add_progress.ensure_number_file_writable(self.path)
        finally:
            blocker.Close()
        add_progress.ensure_number_file_writable(self.path)
        self.assertEqual(before, (self.path.read_bytes(), self.path.stat().st_mtime_ns))

    def test_failed_workbook_save_keeps_record_then_recovers_current_rows(self):
        journal = self.journal()
        journal.record('123456', [1], 'processed')
        with self.assertRaisesRegex(PermissionError, 'Excel lock'):
            journal.recover([('123456', [8]), ('654321', [2])], Mock(side_effect=PermissionError('Excel lock')))
        self.assertTrue(journal.path.exists())
        later = self.journal()
        mark = Mock()
        result = later.recover([('654321', [2]), ('123456', [8, 10])], mark)
        mark.assert_called_once_with([8, 10])
        self.assertEqual(result, [('654321', [2])])
        self.assertFalse(journal.path.exists())

    def test_sending_record_blocks_replay_and_retains_journal(self):
        journal = self.journal()
        journal.record('123456', [1], 'sending')
        mark = Mock()
        with self.assertRaisesRegex(RuntimeError, '123456.*不会自动重复发送') as caught:
            self.journal().recover([('123456', [2])], mark)
        self.assertIn(str(journal.path), str(caught.exception))
        mark.assert_not_called()
        self.assertTrue(journal.path.exists())

    def test_previously_marked_number_clears_uncertain_record_without_writing(self):
        journal = self.journal()
        journal.record('123456', [1], 'sending')
        mark = Mock()
        remaining = [('654321', [2])]
        self.assertEqual(journal.recover(remaining, mark, {'123456', '654321'}), remaining)
        mark.assert_not_called()
        self.assertFalse(journal.path.exists())

    def test_missing_number_is_not_assumed_marked_and_never_uses_old_row(self):
        journal = self.journal()
        journal.record('123456', [1], 'processed')
        mark = Mock()
        with self.assertRaisesRegex(RuntimeError, '不在当前工作表'):
            journal.recover([('654321', [1])], mark, {'654321'})
        mark.assert_not_called()
        self.assertTrue(journal.path.exists())

    def test_invalid_json_stops_without_overwriting(self):
        journal = self.journal()
        journal.path.write_text('{broken', encoding='utf-8')
        with self.assertRaisesRegex(RuntimeError, '恢复记录无法安全读取'):
            self.journal()
        self.assertEqual(journal.path.read_text(encoding='utf-8'), '{broken')

    def test_wrong_sheet_with_pending_record_stops(self):
        self.journal().record('123456', [1], 'processed')
        with self.assertRaisesRegex(RuntimeError, '当前选择的是'):
            self.journal('OtherSheet')

    def test_empty_journal_allows_sheet_change(self):
        journal = self.journal()
        journal._persist({})
        next_sheet = self.journal('OtherSheet')
        next_sheet.record('123456', [4], 'sending')
        self.assertEqual(json.loads(next_sheet.path.read_text(encoding='utf-8'))['sheet'], 'OtherSheet')

    def test_wrong_source_and_invalid_record_stop(self):
        for field, value in [('source', 'another.xlsx'), ('version', 7),
                             ('records', {'123456': {'rows': [0], 'state': 'processed'}})]:
            with self.subTest(field=field):
                journal = self.journal()
                journal.record('123456', [1], 'processed')
                payload = json.loads(journal.path.read_text(encoding='utf-8'))
                payload[field] = value
                journal.path.write_text(json.dumps(payload), encoding='utf-8')
                with self.assertRaisesRegex(RuntimeError, '恢复记录无法安全读取'):
                    self.journal()
                journal.path.unlink()

    def test_clear_failure_after_workbook_save_is_idempotent_on_restart(self):
        journal = self.journal()
        journal.record('123456', [1], 'processed')
        mark = Mock()
        with patch.object(Path, 'unlink', side_effect=PermissionError('blocked')):
            with self.assertRaisesRegex(RuntimeError, '恢复记录暂时无法清理'):
                journal.recover([('123456', [8])], mark)
        mark.assert_called_once_with([8])
        self.assertTrue(journal.path.exists())
        second_mark = Mock()
        self.assertEqual(self.journal().recover([], second_mark, {'123456'}), [])
        second_mark.assert_not_called()
        self.assertFalse(journal.path.exists())

    def test_record_write_failure_preserves_previous_durable_state(self):
        journal = self.journal()
        journal.record('123456', [1], 'sending')
        with patch.object(add_progress.os, 'replace', side_effect=PermissionError('blocked')):
            with self.assertRaisesRegex(RuntimeError, '无法保存好友处理恢复记录'):
                journal.record('123456', [1], 'processed')
        self.assertEqual(journal.records['123456']['state'], 'sending')
        self.assertEqual(self.journal().records['123456']['state'], 'sending')
        self.assertEqual(list(journal.path.parent.glob('*.tmp')), [])

    def test_fsync_failure_does_not_record_or_claim_persistence(self):
        journal = self.journal()
        with patch.object(add_progress.os, 'fsync', side_effect=OSError('disk full')):
            with self.assertRaisesRegex(RuntimeError, 'disk full'):
                journal.record('123456', [1], 'sending')
        self.assertFalse(journal.path.exists())
        self.assertEqual(journal.records, {})

    def test_multiple_records_clear_one_preserves_other(self):
        journal = self.journal()
        journal.record('123456', [1], 'processed')
        journal.record('654321', [2], 'sending')
        journal.clear('123456')
        self.assertEqual(set(self.journal().records), {'654321'})

    def test_uncertain_record_is_checked_before_other_record_recovery(self):
        journal = self.journal()
        journal.record('123456', [1], 'processed')
        journal.record('654321', [2], 'sending')
        mark = Mock()
        with self.assertRaisesRegex(RuntimeError, '不会自动重复发送'):
            journal.recover([('123456', [9]), ('654321', [10])], mark)
        mark.assert_not_called()

    def test_known_numbers_must_include_pending_numbers(self):
        with self.assertRaisesRegex(ValueError, 'all_numbers'):
            self.journal().recover([('123456', [1])], Mock(), set())

    def test_processed_record_cannot_be_downgraded_to_sending(self):
        journal = self.journal()
        journal.record('123456', [1], 'processed')
        with self.assertRaisesRegex(RuntimeError, '请先恢复号码表'):
            journal.record('123456', [1], 'sending')
        self.assertEqual(self.journal().records['123456']['state'], 'processed')


if __name__ == '__main__':
    unittest.main()
