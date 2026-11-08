import json
import tempfile
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import Mock, patch
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import zq_qq_001 as c


class ExistingWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'old.xlsx'
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.stack.enter_context(patch.object(c, 'NUMBER_WORKBOOK_PATH', self.path))
        self.stack.enter_context(patch.object(c, 'NUMBER_JOURNAL_PATH', self.path.with_suffix('.jsonl')))

    def test_existing_first_sheet_is_reused_across_two_tasks(self):
        b = Workbook();s=b.active;s.title='原群旧名称'
        s.append(['群甲','123456']);s.append(['11111111',None,'已处理'])
        s['C2'].font=Font(color='FF123456');b.save(self.path);b.close()
        for new_number in ['22222222','33333333']:
            b,s,n,completed,records=c.open_number_workbook(Workbook,load_workbook)
            c.reconcile_journal(b,s,records)
            g=c.append_group_info(b,s,'改名后的同一群','123456')
            self.assertFalse(c.append_number(b,s,n,'11111111',g))
            self.assertTrue(c.append_number(b,s,n,new_number,g))
            c.reconcile_journal(b,s,c.load_journal_records(),force_save=True);b.close()
        b=load_workbook(self.path)
        self.assertEqual(b.sheetnames,['原群旧名称'])
        self.assertEqual([r[0] for r in b.active.iter_rows(min_row=2,values_only=True)],['11111111','22222222','33333333'])
        self.assertEqual(b.active['C2'].value,'已处理')
        self.assertEqual(b.active['C2'].font.color.rgb,'FF123456')
        self.assertEqual(completed,set());b.close()

    def test_same_number_in_two_groups_and_recovery_are_idempotent(self):
        b,s,n,_,_=c.open_number_workbook(Workbook,load_workbook)
        for gid in ['123456','234567']:
            g=c.append_group_info(b,s,'同名群',gid)
            self.assertTrue(c.append_number(b,s,n,'11111111',g))
        # Simulate a process ending after journal write but before final Excel flush.
        b.close()
        b,s,n,_,records=c.open_number_workbook(Workbook,load_workbook)
        self.assertTrue(c.reconcile_journal(b,s,records))
        self.assertFalse(c.reconcile_journal(b,s,records))
        self.assertEqual(len(b.worksheets),2)
        for sheet in b:self.assertEqual(sheet['A2'].value,'11111111')
        b.close()

    def test_neighbouring_legacy_journal_is_not_imported(self):
        other=self.path.with_name('qq_member_numbers.jsonl')
        other.write_text(json.dumps({'type':'group','group_name':'无关群','group_number':'999999'}),encoding='utf-8')
        c.configure_number_storage(self.path)
        b,_,_,_,records=c.open_number_workbook(Workbook,load_workbook)
        self.assertEqual(records,[])
        self.assertEqual(c.NUMBER_JOURNAL_PATH,self.path.with_suffix('.jsonl'))
        b.close()

    def test_mixed_legacy_rows_are_backed_up_and_preserved(self):
        b=Workbook();s=b.active;s.title='旧表'
        s.append(['甲','123456']);s.append(['11111111',None,'已处理'])
        s.append(['乙','234567']);s.append(['22222222',None,'备注'])
        b.save(self.path);b.close()
        before=self.path.read_bytes()
        b,s,n,_,_=c.open_number_workbook(Workbook,load_workbook)
        self.assertEqual(list(b['历史汇总'].values), [('甲','123456',None),('11111111',None,'已处理'),('乙','234567',None),('22222222',None,'备注')])
        self.assertEqual(b['甲_123456']['C2'].value,'已处理')
        self.assertEqual(b['乙_234567']['C2'].value,'备注')
        backups=list(self.path.parent.glob('*.before_group_append_*.xlsx'))
        self.assertEqual(len(backups),1)
        self.assertEqual(backups[0].read_bytes(),before)
        b.close()

    def test_completed_group_is_not_reset_and_interrupted_group_resumes(self):
        self.assertEqual(c.resume_for_new_task({'completed':True}), {'completed':True})
        self.assertEqual(c.resume_for_new_task({})['scroll_percent'],0.0)
        position={'completed':False,'next_index':3}
        self.assertEqual(c.resume_for_new_task(position),position)

    def test_existing_file_picker_does_not_offer_overwrite(self):
        import desktop_app, gui_config
        field=next(f for f in gui_config.fields('capture') if f.key=='NUMBER_WORKBOOK_PATH')
        app=Mock();app.inputs={field.key:Mock()};app.inputs[field.key].get.return_value=str(self.path)
        with patch('tkinter.filedialog.asksaveasfilename',return_value=str(self.path)) as dialog:
            desktop_app.App.browse(app,field)
        self.assertFalse(dialog.call_args.kwargs['confirmoverwrite'])
        app.inputs[field.key].set.assert_called_once_with(str(self.path))

if __name__=='__main__':unittest.main()

