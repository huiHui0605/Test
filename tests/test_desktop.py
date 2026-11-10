"""Headless GUI checks: never opens WhatsApp or contacts Google."""
import json
import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch
from PySide6.QtWidgets import QApplication,QDialog
from PySide6.QtCore import QEventLoop,QTimer
import main

APP=QApplication.instance() or QApplication([])


def pump(ms=50):
    loop=QEventLoop();QTimer.singleShot(ms,loop.quit);loop.exec()


class DesktopTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
        self.patch1=patch.object(main,'DATA',self.root);self.patch1.start()
        self.patch2=patch.object(main,'SETTINGS',self.root/'settings.json');self.patch2.start()
        self.window=main.Window();self.window.auto_pdf.setChecked(False);self.window.show();pump(260)
    def tearDown(self):
        self.window.dirty=False;self.window.busy_pdf=False;self.window.busy_sync=False
        self.window.save_timer.stop();self.window.watch_timer.stop();self.window.google_timer.stop();self.window.close();self.window.deleteLater();pump()
        self.patch2.stop();self.patch1.stop();self.temp.cleanup()
    def test_interface_and_dialogs_use_english(self):
        import ast,re
        from PySide6.QtWidgets import QWidget,QComboBox,QTableWidget
        tree=ast.parse(Path(main.__file__).read_text(encoding='utf-8'))
        aliases=next(n for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='ALIASES' for t in n.targets))
        allowed={id(n) for n in ast.walk(aliases)}
        # Includes dialogs, validation failures and background status messages.
        for node in ast.walk(tree):
            if isinstance(node,ast.Constant) and isinstance(node.value,str) and id(node) not in allowed:
                self.assertIsNone(re.search(r'[\u3400-\u9fff]',node.value),node.value)
        dialogs=[main.MappingDialog(self.window.headers,{},self.window),main.TemplateDialog(self.window.cfg,self.window),main.GoogleDialog({},self.window)]
        try:
            for parent in [self.window]+dialogs:
                for widget in [parent]+parent.findChildren(QWidget):
                    for method in ('text','windowTitle','placeholderText','suffix'):
                        if hasattr(widget,method):
                            value=getattr(widget,method)()
                            self.assertIsNone(re.search(r'[\u3400-\u9fff]',value),value)
                    if isinstance(widget,QComboBox):
                        for i in range(widget.count()):self.assertIsNone(re.search(r'[\u3400-\u9fff]',widget.itemText(i)))
                    if isinstance(widget,QTableWidget):
                        for i in range(widget.columnCount()):
                            header=widget.horizontalHeaderItem(i)
                            if header:self.assertIsNone(re.search(r'[\u3400-\u9fff]',header.text()))
        finally:
            for dialog in dialogs:dialog.close();dialog.deleteLater()
    def test_english_filter_keeps_multilingual_records(self):
        w=self.window;headers,_=w.store.table('Records')
        w.store.save_rows('Records',headers,[{'Participant Name':'王小明','Type':'志愿者'},{'Participant Name':'Example Family','Type':'beneficiary'}])
        w.load_table('Records');w.category.setCurrentText('volunteer')
        self.assertEqual(len(w.visible),1);self.assertEqual(w.table.item(0,0).text(),'王小明')
        w.category.setCurrentText('All types');self.assertEqual(len(w.visible),1)
        w.workflow_tabs.setCurrentIndex(main.FAMILY_TAB);self.assertEqual(len(w.visible),1);self.assertEqual(w.table.item(0,0).text(),'Example Family')
    def test_simple_add_person_form_and_action_menus(self):
        w=self.window
        def fill_form():
            dialog=QApplication.activeModalWidget()
            dialog.findChild(main.QLineEdit,'person_Participant Name').setText('New Volunteer')
            dialog.findChild(main.QComboBox,'person_Type').setCurrentText('volunteer')
            dialog.accept()
        QTimer.singleShot(0,fill_form);w.add_person_dialog()
        self.assertFalse(w.dirty)
        self.assertEqual(w.store.table('Records')[1][0]['Participant Name'],'New Volunteer')
        self.assertIn('Certificate Templates',[a.text() for a in w.settings_menu.actions()])
        self.assertIn('Delete Selected Rows',[a.text() for a in w.actions_menu.actions()])
        self.assertIn('Generate / Retry PDFs',[a.text() for a in w.actions_menu.actions()])

    def test_three_tabs_and_continuous_blank_rows(self):
        w=self.window;self.assertEqual(w.workflow_tabs.count(),3)
        self.assertEqual(w.table.rowCount(),6)
        w.table.item(4,0).setText('Manual Volunteer')
        self.assertEqual(w.table.rowCount(),11)
        self.assertTrue(w.save())
        self.assertEqual(len(w.rows),1);self.assertEqual(w.rows[0]['Type'],'volunteer')
        self.assertEqual(w.table.rowCount(),7)
        w.workflow_tabs.setCurrentIndex(1);self.assertEqual(len(w.visible),0)
        w.table.item(0,0).setText('Manual Family');w.save()
        self.assertEqual(len(w.store.table('Records')[1]),2)
        self.assertEqual(w.table.item(0,0).text(),'Manual Family')
        w.workflow_tabs.setCurrentIndex(main.ACTIVITY_TAB);self.assertEqual(w.table.item(0,0).text(),'Manual Volunteer')
    def test_family_pdf_and_activity_share_preview(self):
        w=self.window;h,_=w.store.table('Records')
        w.store.save_rows('Records',h,[{'Participant Name':'Family Example','Type':'beneficiary','Phone':'0123456789'},{'Participant Name':'Volunteer Example','Type':'volunteer','Phone':'0123456789'}])
        w.load_table('Records');w.workflow_tabs.setCurrentIndex(main.FAMILY_TAB)
        with patch.object(w,'review_message',return_value=None),patch('main.webbrowser.open') as browser:
            w.share(w.visible[0]);browser.assert_not_called()
        self.assertEqual(len(list((self.root/'family_confirmations').glob('*.pdf'))),1)
        w.workflow_tabs.setCurrentIndex(main.ACTIVITY_TAB);w.cfg['event_message']={'event':'Food Packing','date':'Saturday','time':'09:00','location':'Main Hall','details':'Bring water.','template':main.DEFAULT_MESSAGE}
        with patch.object(w,'review_message',side_effect=lambda message,pdf:message) as preview,patch('main.webbrowser.open') as browser:
            w.share(w.visible[0]);message=preview.call_args.args[0]
            self.assertIn('Volunteer Example',message);self.assertIn('Main Hall',message);self.assertIn('Bring water.',message)
            self.assertTrue(browser.call_args.args[0].startswith('https://wa.me/60123456789?'))

    def test_add_autosave_search_reload(self):
        w=self.window;w.add_row();w.table.item(0,0).setText('Alice');w.table.item(0,1).setText('volunteer')
        pump(1600)
        self.assertFalse(w.dirty);self.assertTrue(w.rows[0]['__fb_id'])
        w.search.setText('Nobody');self.assertEqual(len(w.visible),0)
        w.search.clear();self.assertEqual(len(w.visible),1)
        w.reload();self.assertEqual(w.rows[0]['Participant Name'],'Alice')
    def test_paste_multiple_rows(self):
        w=self.window
        QApplication.clipboard().setText('Alice\tvolunteer\t0123\nBob\tdonor\t0456')
        w.paste_cells();self.assertEqual(len(w.rows),2);self.assertEqual(w.rows[1]['Type'],'donor')
        self.assertTrue(w.save());self.assertEqual(len(w.store.table('Records')[1]),2)
    @unittest.skipUnless(os.environ.get('FOODBANK_TEST_SOFFICE'),'real converter not configured')
    def test_automatic_real_certificates_and_share(self):
        import time
        w=self.window
        h,_=w.store.table('Records')
        w.store.save_rows('Records',h,[{'Participant Name':'Example Volunteer','Type':'volunteer','Date':'2026-11-02'}, {'Participant Name':'Example Donor','Type':'donor','Date':'2026-11-02'}, {'Participant Name':'Example Family','Type':'beneficiary'}])
        w.load_table('Records');w.cfg.update(engine='libreoffice',soffice=os.environ['FOODBANK_TEST_SOFFICE'],auto_pdf=True)
        w.cfg['templates']={'volunteer':{'path':str(main.ROOT/'templates/volunteer_template.pptx'),'marker':'auto'}}
        w.auto_generate()
        deadline=time.monotonic()+30
        while w.busy_pdf and time.monotonic()<deadline:pump(100)
        self.assertFalse(w.busy_pdf)
        for row in w.rows[:1]:
            rec=w.record(row);entry=w.cache[w.cache_key(rec['id'])]
            self.assertEqual(entry['status'],'PDF ready',entry)
            import subprocess
            extracted=subprocess.check_output(['pdftotext',entry['pdf'],'-'],text=True)
            self.assertIn(rec['name'],extracted)
        self.assertEqual(len(list((self.root/'certificates').rglob('*.pdf'))),1)
        with patch('main.webbrowser.open') as browser,patch('main.QDesktopServices.openUrl') as folder:
            w.share(0);self.assertEqual(browser.call_count,1);self.assertEqual(folder.call_count,1)
            self.assertTrue(QApplication.clipboard().text().endswith('.pdf'))
    def test_beneficiary_share_user_action_only(self):
        w=self.window;w.add_row();w.table.item(0,0).setText('Family');w.table.item(0,1).setText('beneficiary');w.table.item(0,2).setText('0123456789');w.save()
        with patch('main.webbrowser.open') as browser:
            w.share(0);self.assertEqual(browser.call_count,1);self.assertTrue(browser.call_args.args[0].startswith('https://wa.me/60'))
    def test_dialogs_and_mapping(self):
        w=self.window;mapping=main.MappingDialog(w.headers,{},w)
        self.assertEqual(mapping.value()['name'],'Participant Name');mapping.close()
        template=main.TemplateDialog(w.cfg,w);self.assertIn('volunteer',template.value()['templates']);template.close()
        google=main.GoogleDialog({},w);google.add({'name':'Full Name'});self.assertEqual(google.value()['sources'][0]['name'],'Full Name');google.close()
    def test_sync_callback_no_duplicate(self):
        w=self.window
        cfg={'sources':[{'destination':'Google Volunteers'}]};w.cfg['google']=cfg
        batches=[{'sheet':'Google Volunteers','source':'abc','headers':['Participant Name','Type'],'rows':[{'Participant Name':'Alice','Type':'volunteer','__fb_source':'abc','__fb_key':'1'}]}]
        with patch('main.fetch',return_value=batches):
            w.sync_google();pump(150);self.assertFalse(w.busy_sync)
            w.sync_google();pump(150);self.assertEqual(len(w.store.table('Google Volunteers')[1]),1)
        w.change_sheet('Google Volunteers');self.assertTrue(w.sheet_is_synced())
        self.assertFalse(w.table.item(0,0).flags() & main.Qt.ItemFlag.ItemIsEditable)
    def test_auto_pdf_all_worksheets_in_background(self):
        w=self.window;w.cfg['auto_pdf']=True
        b={'sheet':'Google Volunteers','source':'abc','headers':['Participant Name','Type'],'rows':[{'Participant Name':'Volunteer','Type':'volunteer','__fb_source':'abc','__fb_key':'1'}]}
        w.store.merge_google([b]);w.load_table('Records')
        with patch('main.generate_one',return_value={'status':'PDF ready','pdf':'fake','signature':'s'}) as generate:
            w.auto_generate();pump(150)
            self.assertEqual(generate.call_count,1);self.assertEqual(generate.call_args.args[0]['name'],'Volunteer');self.assertFalse(w.busy_pdf)
    def test_certificate_columns_persist_and_retired_donors_stay_in_excel(self):
        w=self.window;h,_=w.store.table('Records');h+=['Printed Name','Event Day','Reference']
        w.store.save_rows('Records',h,[{'Participant Name':'Volunteer','Type':'volunteer','Printed Name':'Alice','Event Day':'2026-12-01','Reference':'0007'},{'Participant Name':'Historical Donor','Type':'donor'}])
        w.load_table('Records')
        for key,col in [('name','Printed Name'),('date','Event Day'),('cert_no','Reference')]:w.column_boxes[key].setCurrentText(col)
        rec=w.record(w.rows[0]);self.assertEqual((rec['name'],rec['date'],rec['cert_no']),('Alice','2026-12-01','0007'))
        w.load_table('Records');self.assertEqual(w.column_boxes['cert_no'].currentText(),'Reference')
        self.assertEqual(len(w.rows),2);self.assertEqual(len(w.visible),1)
        w.cfg['auto_pdf']=True
        with patch('main.generate_one',return_value={'status':'Test'}) as generate:
            w.auto_generate();pump(150)
            self.assertEqual(generate.call_count,1);self.assertEqual(generate.call_args.args[0]['type'],'volunteer')
        dialog=main.TemplateDialog(w.cfg,w);self.assertNotIn('donor',dialog.paths);dialog.close()
        dialog=main.GoogleDialog({},w);dialog.add({});box=dialog.table.cellWidget(0,0)
        self.assertEqual([box.itemText(i) for i in range(box.count())],['volunteer','beneficiary']);dialog.close()

    def test_existing_workbook_gets_certificate_column_without_losing_ids(self):
        from openpyxl import Workbook
        w=self.window;old=self.root/'legacy.xlsx';book=Workbook();sheet=book.active;sheet.title='Records'
        sheet.append(['Participant Name','Type','Phone','Email','Date','__fb_id']);sheet.append(['Alice','volunteer','','','2026-12-01','stable-id']);book.save(old);book.close()
        w.store=main.ExcelStore(old);w.load_table('Records')
        self.assertIn('Certificate Number',w.headers)
        w.table.item(0,w.headers.index('Certificate Number')).setText('0007');self.assertTrue(w.save())
        row=w.store.table('Records')[1][0];self.assertEqual(row['__fb_id'],'stable-id');self.assertEqual(row['Certificate Number'],'0007')
        self.assertEqual(w.column_boxes['cert_no'].currentText(),'Certificate Number')

    def test_external_change_during_edit_not_overwritten(self):
        w=self.window;w.add_row();w.table.item(0,0).setText('Local')
        other=main.ExcelStore(w.store.path);h,_=other.table('Records');other.save_rows('Records',h,[{'Participant Name':'External'}])
        w.watch();self.assertTrue(w.dirty);self.assertFalse(w.save_timer.isActive())
        self.assertEqual(main.ExcelStore(w.store.path).table('Records')[1][0]['Participant Name'],'External')


if __name__=='__main__':unittest.main()
