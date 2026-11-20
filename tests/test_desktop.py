"""Headless GUI checks: never opens WhatsApp or contacts Google."""
import json
import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import tempfile
from pathlib import Path
import unittest
from datetime import datetime
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
        self.identity_patch=patch.object(main.GoogleAccount,'identity',side_effect=RuntimeError('Disconnected test account'));self.identity_patch.start()
        self.window=main.Window();self.window.auto_pdf.setChecked(False);self.window.show();pump(260)
    def tearDown(self):
        self.window.dirty=False;self.window.busy_pdf=False;self.window.busy_sync=False
        self.window.save_timer.stop();self.window.watch_timer.stop();self.window.google_timer.stop();self.window.close();self.window.deleteLater();pump()
        self.identity_patch.stop();self.patch2.stop();self.patch1.stop();self.temp.cleanup()
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
        w.store.save_rows('Records',h,[{'Participant Name':'Family Example','Type':'beneficiary','Phone':'0123456789','Email':'family@example.test'},{'Participant Name':'Volunteer Example','Type':'volunteer','Phone':'0123456789','Date':'2026-12-01'}])
        w.load_table('Records');w.workflow_tabs.setCurrentIndex(main.FAMILY_TAB)
        with patch.object(w,'review_email',return_value=None),patch('main.webbrowser.open') as browser:
            w.share(w.visible[0]);browser.assert_not_called()
        self.assertEqual(len(list((self.root/'family_confirmations').glob('*.pdf'))),1)
        w.cfg['mapping'][w.mapping_key()]={'activity_date':'Date'}
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
        h+=['Printed Name','Attend Date','Cert No']
        people=[('Alice Example','2026-11-20','0007'),('Bob Example','2026-11-21','0008'),('Alice Example','2026-11-22','0009')]
        w.store.save_rows('Records',h,[{'Participant Name':'Unmapped Name','Type':'volunteer','Date':'WRONG DATE','Certificate Number':'WRONG NUMBER',
            'Printed Name':name,'Attend Date':datetime.fromisoformat(date),'Cert No':number} for name,date,number in people]+
            [{'Participant Name':'Example Donor','Type':'donor'},{'Participant Name':'Example Family','Type':'beneficiary'}])
        w.load_table('Records');w.cfg.update(engine='libreoffice',soffice=os.environ['FOODBANK_TEST_SOFFICE'],auto_pdf=True)
        w.cfg['mapping'][w.mapping_key()]={'name':'Printed Name','date':'Attend Date','cert_no':'Cert No'}
        w.cfg['templates']={'volunteer':{'path':str(main.ROOT/'templates/volunteer_template.pptx'),'marker':'auto'}}
        w.auto_generate()
        deadline=time.monotonic()+60
        while w.busy_pdf and time.monotonic()<deadline:pump(100)
        self.assertFalse(w.busy_pdf)
        for row,expected in zip(w.rows[:3],people):
            rec=w.record(row);entry=w.cache[w.cache_key(rec['id'])]
            self.assertEqual(entry['status'],'PDF ready',entry)
            import subprocess
            extracted=subprocess.check_output(['pdftotext',entry['pdf'],'-'],text=True)
            for value in expected:self.assertIn(value,extracted)
            self.assertNotIn('T00:00:00',extracted)
            for value in ('July 18, 2026','26/06/3705','WRONG DATE','WRONG NUMBER','Unmapped Name'):
                self.assertNotIn(value,extracted)
            for other in people:
                if other!=expected:
                    self.assertNotIn(other[1],extracted);self.assertNotIn(other[2],extracted)
        self.assertEqual(len(list((self.root/'certificates').rglob('*.pdf'))),3)
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
        google=main.GoogleDialog({},w);self.assertEqual(len(google.panels),2);self.assertEqual(google.panels[0].kind,'volunteer');google.close()

    def test_template_dialog_rejects_fixed_date_and_number(self):
        from pptx import Presentation
        from pptx.util import Inches
        path=self.root/'fixed.pptx';prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(0,0,Inches(6),Inches(2)).text='Name / July 18, 2026 / 26/06/3705';prs.save(path)
        dialog=main.TemplateDialog({'templates':{'volunteer':{'path':str(path)}}},self.window)
        with patch('main.QMessageBox.warning') as warning,patch.object(dialog,'accept') as accept:
            dialog.validate();accept.assert_not_called();self.assertIn('Template missing Date No',warning.call_args.args[2])
        dialog.close()

    def test_attend_date_and_cert_no_column_aliases(self):
        mapping=main.MappingDialog(['Name','Attend Date','Cert No'],{},self.window)
        self.assertEqual(mapping.value()['date'],'Attend Date')
        self.assertEqual(mapping.value()['cert_no'],'Cert No');mapping.close()

    def test_date_display_and_invalid_number_do_not_rewrite_excel(self):
        w=self.window;headers,_=w.store.table('Records')
        source_date=datetime(2026,11,20);wrong_number=datetime(3705,6,26)
        w.store.save_rows('Records',headers,[{'Participant Name':'Alice','Type':'volunteer',
            'Date':source_date,'Certificate Number':wrong_number}])
        w.load_table('Records');rec=w.record(w.rows[0])
        self.assertEqual(rec['date'],'2026-11-20');self.assertTrue(rec['cert_no_is_date'])
        self.assertEqual(rec['cert_no'],'')
        self.assertEqual(w.table.item(0,w.headers.index('Date')).text(),'2026-11-20')
        self.assertNotIn('T00:00:00',w.table.item(0,w.headers.index('Certificate Number')).text())
        w.table.item(0,w.headers.index('Phone')).setText('0123456789');self.assertTrue(w.save())
        row=w.store.table('Records')[1][0]
        self.assertEqual(row['Date'],source_date);self.assertEqual(row['Certificate Number'],wrong_number)
        w.table.item(0,w.headers.index('Certificate Number')).setText('26/06/3705');self.assertTrue(w.save())
        rec=w.record(w.rows[0]);self.assertEqual(rec['cert_no'],'26/06/3705')
        self.assertFalse(rec.get('cert_no_is_date',False))
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
        dialog=main.GoogleDialog({},w)
        self.assertEqual([panel.kind for panel in dialog.panels],['volunteer','beneficiary']);dialog.close()

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

    def prepare_family(self):
        w=self.window;headers,_=w.store.table('Records')
        w.store.save_rows('Records',headers,[{'Participant Name':'Example Family','Type':'beneficiary','Email':'family@example.test','Phone':'0123456789'}])
        w.load_table('Records');w.workflow_tabs.setCurrentIndex(main.FAMILY_TAB)
        return w,w.visible[0]

    def test_email_cancel_never_reserves_or_sends(self):
        w,index=self.prepare_family()
        with patch.object(w,'review_email',return_value=None),patch('main.send_reviewed') as send:
            w.share(index);send.assert_not_called()
        key=w.cache_key(w.record(w.rows[index])['id'])
        self.assertEqual(w.email_journal.get(key),{});self.assertFalse(w.email_busy)
        self.assertEqual(w.table.cellWidget(0,len(w.headers)+1).text(),'Review & Email')

    def test_email_double_click_and_retry_require_new_review(self):
        w,index=self.prepare_family();callbacks=[]
        def queued(fn,complete,failed):callbacks.append((fn,complete))
        with patch.object(w,'review_email',return_value={'raw':'fictional'}) as review,patch.object(w,'submit',side_effect=queued):
            w.share(index);w.share(index)
            self.assertEqual(review.call_count,1);self.assertEqual(len(callbacks),1)
        key=next(iter(w.email_busy));self.assertEqual(w.email_journal.get(key)['status'],'sending')
        w.email_journal.finish(key,'unknown');callbacks[0][1]({'status':'unknown'})
        self.assertFalse(w.email_busy);self.assertIn('check Gmail Sent',w.info.text())
        with patch.object(w,'review_email',return_value=None) as review:
            w.share(index)
            self.assertEqual(review.call_args.args[3]['status'],'unknown')
        self.assertEqual(w.email_journal.get(key)['status'],'unknown')

    def test_review_dialog_requires_review_and_explicit_resend(self):
        from PySide6.QtWidgets import QDialogButtonBox,QCheckBox
        w,index=self.prepare_family()
        def inspect():
            dialog=QApplication.activeModalWidget();buttons=dialog.findChild(QDialogButtonBox)
            send=buttons.button(QDialogButtonBox.StandardButton.Ok)
            self.assertFalse(send.isEnabled())
            boxes=dialog.findChildren(QCheckBox)
            reviewed=next(b for b in boxes if b.text().startswith('I have reviewed'))
            repeat=next(b for b in boxes if b.text().startswith('I checked Gmail'))
            reviewed.setChecked(True);self.assertFalse(send.isEnabled())
            repeat.setChecked(True);self.assertTrue(send.isEnabled());dialog.reject()
        QTimer.singleShot(0,inspect)
        self.assertIsNone(w.review_email('family@example.test','Example Family',self.root/'unused.pdf',{'status':'sent'}))

    def test_email_missing_address_does_not_open_browser_or_send(self):
        w,index=self.prepare_family();w.rows[index]['Email']=''
        w.store.save_rows(w.sheet,w.headers,w.rows);w.load_table(w.sheet)
        with patch.object(w,'fail') as fail,patch.object(w,'review_email') as preview,patch('main.send_reviewed') as send:
            w.share(index);preview.assert_not_called();send.assert_not_called();self.assertIn('email address',fail.call_args.args[0])

    def test_activity_uses_source_date_even_when_certificate_mapping_differs(self):
        w=self.window
        batch={'sheet':'Google Volunteers','source':'source','headers':['Participant Name','Type','Phone','Date','Timestamp'],'rows':[{'Participant Name':'Example A','Type':'volunteer','Phone':'0123456789','Date':'2026-12-01','Timestamp':'WRONG','__fb_source':'source','__fb_key':'a'}, {'Participant Name':'Example B','Type':'volunteer','Phone':'0123456788','Date':'2026-12-08','Timestamp':'WRONG','__fb_source':'source','__fb_key':'b'}]}
        w.store.merge_google([batch]);w.load_table('Google Volunteers');w.workflow_tabs.setCurrentIndex(main.ACTIVITY_TAB)
        w.cfg['mapping'][w.mapping_key()]={'date':'Timestamp'}
        w.cfg['event_message']={'event':'Packing','date':'GLOBAL WRONG','template':main.DEFAULT_MESSAGE}
        with patch.object(w,'review_message',return_value=None) as review,patch('main.webbrowser.open') as browser:
            w.share(0);self.assertIn('2026-12-01',review.call_args.args[0]);self.assertNotIn('WRONG',review.call_args.args[0])
            w.share(1);self.assertIn('2026-12-08',review.call_args.args[0]);browser.assert_not_called()

    def test_source_picker_metadata_headers_and_saved_mapping(self):
        from unittest.mock import Mock
        w=self.window;account=Mock()
        account.spreadsheets.return_value=[{'id':'volunteer-sheet','name':'Tally Responses'},{'id':'family-sheet','name':'Family Responses'}]
        account.worksheets.return_value=[{'sheetId':7,'title':'Responses'}]
        account.headers.return_value=['Full Name','Phone','Email','Participation date','Submission ID']
        def immediately(fn,done,failed):done(fn())
        with patch.object(w,'submit',side_effect=immediately):
            dialog=main.GoogleDialog({},w,account=account);dialog.list_sheets()
            for index,panel in enumerate(dialog.panels):
                panel.sheet.setCurrentIndex(index+1);panel.worksheet.setCurrentIndex(1)
                for key,header in [('name','Full Name'),('phone','Phone'),('email','Email'),('date','Participation date'),('key','Submission ID')]:panel.fields[key].setCurrentText(header)
            result=dialog.value();self.assertEqual(len(result['sources']),2)
            self.assertEqual(result['sources'][0]['date'],'Participation date');self.assertEqual(result['sources'][1]['email'],'Email')
            self.assertNotIn('credentials',result);self.assertEqual(account.headers.call_count,2)
            dialog.close()
            restored=main.GoogleDialog(result,w,account=account);restored.list_sheets()
            self.assertEqual(restored.value(),result);restored.close()

    def test_source_picker_cancel_keeps_settings_and_rejects_id_change(self):
        w=self.window;original={'sources':[{'type':'volunteer','spreadsheet':'sheet','worksheet_id':7,'destination':'Volunteers','key':'Old ID'}]}
        dialog=main.GoogleDialog(original,w);panel=dialog.panels[0]
        panel.sheet.addItem('Sheet','sheet');panel.sheet.blockSignals(True);panel.sheet.setCurrentIndex(1);panel.sheet.blockSignals(False)
        panel.worksheet.blockSignals(True);panel.worksheet.addItem('Responses',{'sheetId':7,'title':'Responses'});panel.worksheet.setCurrentIndex(1);panel.worksheet.blockSignals(False)
        for key,value in [('name','Name'),('phone','Phone'),('date','Participation date'),('key','New ID')]:panel.fields[key].addItem(value,value);panel.fields[key].setCurrentIndex(1)
        with self.assertRaisesRegex(ValueError,'new local worksheet'):panel.value()
        dialog.reject();self.assertEqual(original['sources'][0]['key'],'Old ID')

    def test_google_network_work_does_not_block_event_loop(self):
        import threading
        w=self.window;release=threading.Event();completed=[];ticks=[]
        def network_wait():release.wait(2);return 'staff@example.test'
        timer=QTimer();timer.setInterval(10);timer.timeout.connect(lambda:ticks.append(True));timer.start()
        w.google_task(network_wait,lambda result:completed.append(result))
        pump(80);self.assertTrue(w.busy_google);self.assertGreater(len(ticks),2)
        release.set();pump(100);timer.stop()
        self.assertEqual(completed,['staff@example.test']);self.assertFalse(w.busy_google)

    def test_disconnect_clears_authorization_and_stops_auto_sync(self):
        w=self.window;w.cfg['google']['enabled']=True;w.setup_google_timer()
        with patch.object(w.google_account,'disconnect') as disconnect:
            w.disconnect_google();pump(100);disconnect.assert_called_once()
        self.assertFalse(w.cfg['google']['enabled']);self.assertFalse(w.google_timer.isActive())
        self.assertTrue(w.store.path.exists());self.assertIn('disconnected',w.google_status.text())

    def test_email_attempt_survives_window_restart(self):
        w,index=self.prepare_family();key=w.cache_key(w.record(w.rows[index])['id'])
        w.email_journal.reserve(key)
        from email_delivery import DeliveryJournal
        self.assertIn('uncertain',DeliveryJournal(self.root/'email_delivery.sqlite3').label(key))
        with patch.object(w,'review_email',return_value=None) as review,patch('main.send_reviewed') as send:
            w.share(index);self.assertEqual(review.call_args.args[3]['status'],'sending');send.assert_not_called()


if __name__=='__main__':unittest.main()
