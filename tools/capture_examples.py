"""Capture real Qt widgets using fictitious data; never connect Google or send messages."""
import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import sys
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer,QEventLoop
from applications import family_pdf
from certificates import generate_one


def pump(ms=50):
    loop=QEventLoop();QTimer.singleShot(ms,loop.quit);loop.exec()


def capture():
    app=QApplication.instance() or QApplication([])
    with tempfile.TemporaryDirectory(prefix='foodbank-fictional-demo-') as tmp:
        os.environ['FOODBANK_DATA_DIR']=tmp
        import main
        with patch.object(main.GoogleAccount,'identity',side_effect=RuntimeError('Demo only — no live Google account')):
            w=main.Window();w.auto_pdf.setChecked(False);w.resize(1380,850)
            headers,_=w.store.table('Records')
            w.store.save_rows('Records',headers,[{'Participant Name':'Example Volunteer','Type':'volunteer','Date':'2026-12-01','Certificate Number':'0007'}, {'Participant Name':'Example Family','Type':'beneficiary','Email':'family@example.test','Date':'2026-12-01'}])
            w.load_table('Records');w.show();pump(650)
            if os.environ.get('FOODBANK_TEST_SOFFICE'):
                cfg={'volunteer':{'path':str(ROOT/'examples/volunteer_sample.pptx'),'marker':'Name'}}
                w.cfg.update(templates=cfg,engine='libreoffice',soffice=os.environ['FOODBANK_TEST_SOFFICE'])
                rec=w.record(w.rows[0]);entry=generate_one(rec,cfg,Path(tmp)/'certificates',engine='libreoffice',soffice=os.environ['FOODBANK_TEST_SOFFICE'])
                w.cache[w.cache_key(rec['id'])]=entry;w.render()
            w.summary.setText('FICTITIOUS DEMO · Local Excel · 2 records')
            w.google_status.setText('Google: demo only, not connected');w.say('Example records only. No live account or messages.');pump()
            w.grab().save(str(ROOT/'examples/interface.png'))
            account=Mock();account.spreadsheets.return_value=[{'id':'demo-volunteers','name':'DEMO · Tally Volunteer Responses'},{'id':'demo-families','name':'DEMO · Family Application Responses'}]
            account.worksheets.return_value=[{'sheetId':7,'title':'Responses','gridProperties':{'columnCount':20}}]
            account.headers.return_value=['Full Name','Phone','Email','Selected Participation Date','Submission ID']
            with patch.object(w,'submit',side_effect=lambda fn,done,failed:done(fn())):
                dialog=main.GoogleDialog({},w,account=account);dialog.list_sheets()
                for index,panel in enumerate(dialog.panels):
                    panel.sheet.setCurrentIndex(index+1);panel.worksheet.setCurrentIndex(1)
                    for key,header in [('name','Full Name'),('phone','Phone'),('email','Email'),('date','Selected Participation Date'),('key','Submission ID')]:panel.fields[key].setCurrentText(header)
                dialog.show();pump();dialog.grab().save(str(ROOT/'examples/google_selection.png'));dialog.close()
            row={'Google · Full Name':'Example Family','Google · Identification':'DEMO-IC-NOT-REAL', 'Google · Email':'family@example.test','Google · Household members':'4', 'Google · Address':'Fictitious Example Street', 'Google · Notes':'Please review all details.\nThis is a demonstration only.', 'Google · Other answer':''}
            pdf=family_pdf(row,list(row),'Example Family','demo-only',Path(tmp)/'family_confirmations')
            import shutil
            shutil.copyfile(pdf,ROOT/'examples/family_confirmation_demo.pdf')
            w.workflow_tabs.setCurrentIndex(main.FAMILY_TAB);w.google_identity='DEMO ONLY · staff@example.test'
            def screenshot_review():
                dialog=app.activeModalWidget();dialog.grab().save(str(ROOT/'examples/family_email_review.png'));dialog.reject()
            QTimer.singleShot(150,screenshot_review)
            with patch('main.QDesktopServices.openUrl'):
                w.review_email('family@example.test','Example Family',pdf,{})
            w.save_timer.stop();w.watch_timer.stop();w.google_timer.stop();w.close();w.deleteLater();pump()


if __name__=='__main__':capture()
