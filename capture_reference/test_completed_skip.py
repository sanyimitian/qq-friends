import tempfile
import unittest
from pathlib import Path
import zq_qq_001 as c
from capture_progress import CaptureProgress

class CompletedSkipTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path=Path(self.temp.name)/'existing.xlsx'
        self.progress=CaptureProgress(self.path,'12345')

    def test_completed_group_survives_restart_and_partial_group_resumes(self):
        self.progress.save('111111',group_name='已完成群',completed=True)
        self.progress.save('222222',group_name='未完成群',completed=False,page_names=['甲','乙'],
                           next_index=1,page_turns=2,scroll_percent=20,saved_count=3)
        reopened=CaptureProgress(self.path,'12345')
        self.assertEqual(c.completed_groups_for_resume([],reopened),{'111111'})
        self.assertEqual(c.resume_for_new_task(reopened.get('222222'))['next_index'],1)

    def test_group_metadata_or_saved_members_do_not_imply_completion(self):
        records=[{'type':'group','group_number':'111111','group_name':'旧群'},
                 {'type':'member','group_number':'111111','number':'99999999'}]
        self.assertEqual(c.completed_groups_for_resume(records,self.progress),set())

    def test_account_scoped_journal_can_recover_completion(self):
        records=[{'type':'group_complete','group_number':'111111','account':'12345'},
                 {'type':'group_complete','group_number':'222222','account':'54321'}]
        self.assertEqual(c.completed_groups_for_resume(records,self.progress),{'111111'})

    def test_unfinished_checkpoint_overrides_stale_complete_journal(self):
        self.progress.save('111111',completed=False,next_index=0,page_turns=0,saved_count=0)
        records=[{'type':'group_complete','group_number':'111111','account':'12345'}]
        self.assertEqual(c.completed_groups_for_resume(records,self.progress),set())

    def test_legacy_journal_recovered_only_when_account_is_unambiguous(self):
        records=[{'type':'group_complete','group_number':'111111'}]
        self.assertEqual(c.completed_groups_for_resume(records,self.progress),{'111111'})
        self.progress.data['accounts']['54321']={}
        self.assertEqual(c.completed_groups_for_resume(records,self.progress),set())

    def test_later_restart_records_invalidate_old_completion(self):
        records=[{'type':'group_complete','group_number':'111111'},
                 {'type':'group','group_number':'111111','group_name':'重开'}]
        self.assertEqual(c.completed_groups_for_resume(records,self.progress),set())

    def test_other_workbook_and_account_are_not_skipped(self):
        self.progress.save('111111',completed=True)
        self.assertEqual(c.completed_groups_for_resume([],CaptureProgress(self.path,'54321')),set())
        self.assertEqual(c.completed_groups_for_resume([],CaptureProgress(self.path.with_name('new.xlsx'),'12345')),set())

if __name__=='__main__':unittest.main()

