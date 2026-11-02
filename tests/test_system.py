import copy
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from openpyxl import Workbook,load_workbook
from pptx import Presentation
from pptx.util import Inches
import certificates as cert
from storage import ExcelStore,ConflictError
from google_sync import fetch,spreadsheet_id


def template(path, marker='Name'):
    p=Presentation();s=p.slides.add_slide(p.slide_layouts[6])
    shape=s.shapes.add_textbox(Inches(1),Inches(1),Inches(7),Inches(1))
    para=shape.text_frame.paragraphs[0]
    para.add_run().text='Award to '
    para.add_run().text=marker[:2];para.add_run().text=marker[2:]
    para.add_run().text=' / {{DATE}}'
    p.save(path)


class StorageTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name);self.path=self.root/'db.xlsx';self.store=ExcelStore(self.path)
    def tearDown(self):self.temp.cleanup()
    def test_save_ids_and_backup(self):
        headers,rows=self.store.table('Records');self.store.save_rows('Records',headers,[{'Participant Name':'Alice','Type':'volunteer'}])
        _,rows=self.store.table('Records');uid=rows[0]['__fb_id'];self.assertTrue(uid)
        rows[0]['Participant Name']='Bob';self.store.save_rows('Records',headers,rows)
        self.assertEqual(self.store.table('Records')[1][0]['__fb_id'],uid)
        self.assertEqual(len(list(self.store.backups.glob('*.xlsx'))),2)
    def test_conflict_does_not_overwrite(self):
        w=load_workbook(self.path);w.active['A2']='External';w.save(self.path);w.close()
        with self.assertRaises(ConflictError):self.store.save_rows('Records',[],[])
        self.assertEqual(load_workbook(self.path).active['A2'].value,'External')
        self.assertFalse(self.path.with_name(self.path.name+'.foodbank.lock').exists())
    def test_preserve_formula_and_other_sheets(self):
        w=load_workbook(self.path);w.active.append(['Alice','volunteer',123,None,'=1+1']);w.create_sheet('Other')['A1']='keep';w.save(self.path);w.close();self.store.reload()
        h,r=self.store.table('Records');r[0]['Participant Name']='Bob';self.store.save_rows('Records',h,r)
        w=load_workbook(self.path);self.assertEqual(w['Records']['C2'].value,123);self.assertEqual(w['Records']['E2'].data_type,'f');self.assertEqual(w['Other']['A1'].value,'keep');w.close()
    def test_new_strings_not_formulas(self):
        h,_=self.store.table('Records');self.store.save_rows('Records',h,[{'Participant Name':'=malicious()'}])
        w=load_workbook(self.path);self.assertEqual(w.active['A2'].data_type,'s');w.close()
    def test_columns_added_after_metadata_do_not_overwrite_ids(self):
        h,_=self.store.table('Records');self.store.save_rows('Records',h,[{'Participant Name':'A'}])
        uid=self.store.table('Records')[1][0]['__fb_id']
        w=load_workbook(self.path);ws=w['Records'];ws.cell(1,ws.max_column+1,'Extra');ws.cell(2,ws.max_column,'old');w.save(self.path);w.close();self.store.reload()
        h,r=self.store.table('Records');r[0]['Extra']='new';self.store.save_rows('Records',h,r)
        self.assertEqual(self.store.table('Records')[1][0]['Extra'],'new');self.assertEqual(self.store.table('Records')[1][0]['__fb_id'],uid)
    def test_duplicate_copied_ids_repaired(self):
        h,_=self.store.table('Records');self.store.save_rows('Records',h,[{'Participant Name':'A'},{'Participant Name':'B'}])
        w=load_workbook(self.path);ws=w['Records'];ws.cell(3,6,ws.cell(2,6).value);w.save(self.path);w.close();self.store.reload()
        self.store.ensure_ids('Records');ids=[r['__fb_id'] for r in self.store.table('Records')[1]]
        self.assertEqual(len(set(ids)),2)
    def batch(self):
        return {'sheet':'Google Volunteers','source':'source','headers':['Participant Name','Type'],'rows':[{'Participant Name':'Alice','Type':'volunteer','__fb_key':'id1','__fb_source':'source'}]}
    def test_sync_idempotent_and_update(self):
        b=self.batch();self.assertEqual(self.store.merge_google([b]),(1,0))
        uid=self.store.table(b['sheet'])[1][0]['__fb_id']
        stamp=self.store.stamp();backups=len(list(self.store.backups.glob('*.xlsx')))
        self.assertEqual(self.store.merge_google([b]),(0,0));self.assertEqual(stamp,self.store.stamp());self.assertEqual(backups,len(list(self.store.backups.glob('*.xlsx'))));b['rows'][0]['Participant Name']='Bob'
        self.assertEqual(self.store.merge_google([b]),(0,1));self.assertEqual(self.store.table(b['sheet'])[1][0]['__fb_id'],uid)
    def test_sync_never_takes_over_local_sheet(self):
        b=self.batch();b['sheet']='Records'
        with self.assertRaises(ValueError):self.store.merge_google([b])
        self.assertEqual(self.store.table('Records')[1],[])
    def test_sync_does_not_delete_missing_records(self):
        b=self.batch();self.store.merge_google([b]);b['rows']=[];self.store.merge_google([b]);self.assertEqual(len(self.store.table(b['sheet'])[1]),1)
    def test_new_and_delete_rows(self):
        h,_=self.store.table('Records');self.store.save_rows('Records',h,[{'Participant Name':'A'},{'Participant Name':'B'}])
        _,r=self.store.table('Records');self.store.save_rows('Records',h,[r[1]],deleted=[r[0]['_row']])
        self.assertEqual(self.store.table('Records')[1][0]['Participant Name'],'B')


class CertificateTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name);self.path=self.root/'template.pptx';template(self.path)
    def tearDown(self):self.temp.cleanup()
    def test_split_runs_and_keep_prefix(self):
        out=self.root/'out.pptx';self.assertEqual(cert.fill_template(self.path,out,{'NAME':'王小明','DATE':'2026-11-02'}),1)
        p=Presentation(out);para=p.slides[0].shapes[0].text_frame.paragraphs[0]
        self.assertEqual(para.text,'Award to 王小明 / 2026-11-02');self.assertEqual(para.runs[0].text,'Award to ')
    def test_custom_marker_and_missing(self):
        template(self.path,'Recipient');out=self.root/'out.pptx'
        with self.assertRaises(ValueError):cert.fill_template(self.path,out,{'NAME':'Alice'})
        cert.fill_template(self.path,out,{'NAME':'Alice'},'Recipient')
        self.assertIn('Alice',Presentation(out).slides[0].shapes[0].text)
    def test_table_marker(self):
        p=Presentation();s=p.slides.add_slide(p.slide_layouts[6]);s.shapes.add_table(1,1,0,0,Inches(2),Inches(1)).table.cell(0,0).text='{{NAME}}';p.save(self.path)
        out=self.root/'out.pptx';cert.fill_template(self.path,out,{'NAME':'Alice'})
        self.assertEqual(Presentation(out).slides[0].shapes[0].table.cell(0,0).text,'Alice')
    def test_cache_and_regeneration(self):
        rec={'id':'abc','name':'Alice','type':'donor','date':'2026'};cfg={'donor':{'path':str(self.path)}}
        def convert(pptx,out,*args):Path(out).write_bytes(b'%PDF-1.4\nmock')
        with patch('certificates.convert_pdf',side_effect=convert) as mock:
            first=cert.generate_one(rec,cfg,self.root/'output');second=cert.generate_one(rec,cfg,self.root/'output',first)
            self.assertEqual(mock.call_count,1);self.assertEqual(first,second)
            rec['name']='Bob';third=cert.generate_one(rec,cfg,self.root/'output',first);self.assertEqual(mock.call_count,2);self.assertNotEqual(first['signature'],third['signature'])
    def test_beneficiary_and_missing_template(self):
        self.assertEqual(cert.generate_one({'type':'beneficiary'},{},self.root)['status'],'Not applicable')
        self.assertIn('template required',cert.generate_one({'type':'donor','name':'A'},{},self.root)['status'])
    def test_whatsapp_and_aliases(self):
        self.assertTrue(cert.whatsapp_url('012-3456789','A').startswith('https://wa.me/60123456789?'))
        self.assertEqual(cert.kind('捐赠者'),'donor');self.assertEqual(cert.kind('志愿者'),'volunteer')
    @unittest.skipUnless(os.environ.get('FOODBANK_TEST_SOFFICE'),'Set FOODBANK_TEST_SOFFICE for real PDF integration')
    def test_real_pdf(self):
        rec={'id':'abc','name':'Alice Example','type':'volunteer','date':'2026-11-02'}
        result=cert.generate_one(rec,{'volunteer':{'path':str(self.path)}},self.root/'pdf',engine='libreoffice',soffice=os.environ['FOODBANK_TEST_SOFFICE'])
        self.assertTrue(Path(result['pdf']).read_bytes().startswith(b'%PDF-'))
        import subprocess
        extracted=subprocess.check_output(['pdftotext',result['pdf'],'-'],text=True)
        self.assertIn('Alice Example',extracted);self.assertNotIn('Name',extracted)


class FakeService:
    def __init__(self,grid):self.grid=grid
    def spreadsheets(self):return self
    def values(self):return self
    def get(self,**kwargs):return self
    def execute(self,**kwargs):return {'values':self.grid}


class GoogleTests(unittest.TestCase):
    def cfg(self):return {'sources':[{'type':'volunteer','spreadsheet':'https://docs.google.com/spreadsheets/d/abc_123/edit','range':'Responses!A1:Z','destination':'Google Volunteers','name':'Full Name','phone':'Phone','key':'Response ID'}]}
    def test_read_mapping_and_raw_columns(self):
        batches=fetch(self.cfg(),FakeService([['Full Name','Phone','Response ID'],['Alice','0123','1']]))
        rec=batches[0]['rows'][0];self.assertEqual(rec['Participant Name'],'Alice');self.assertEqual(rec['Google · Phone'],'0123');self.assertEqual(rec['__fb_key'],'1')
    def test_missing_header_and_duplicate_key_rejected(self):
        with self.assertRaises(ValueError):fetch(self.cfg(),FakeService([['Name']]))
        with self.assertRaises(ValueError):fetch(self.cfg(),FakeService([['Full Name','Phone','Response ID'],['A','0','1'],['B','1','1']]))
    def test_append_row_key_offset(self):
        cfg=self.cfg();cfg['sources'][0]['key']='';cfg['sources'][0]['range']='Responses!A5:Z'
        rec=fetch(cfg,FakeService([['Full Name','Phone','Response ID'],['A','0','']]))[0]['rows'][0]
        self.assertEqual(rec['__fb_key'],'row:6')
    def test_safe_destination_and_duplicate_headers(self):
        cfg=self.cfg();cfg['sources'][0]['destination']='bad/name'
        with self.assertRaises(ValueError):fetch(cfg,FakeService([]))
        with self.assertRaises(ValueError):fetch(self.cfg(),FakeService([['Full Name','Phone','Phone']]))
    def test_sheet_url(self):self.assertEqual(spreadsheet_id('https://docs.google.com/spreadsheets/d/abc_123/edit'),'abc_123')


if __name__=='__main__':unittest.main()
