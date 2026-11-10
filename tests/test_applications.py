import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import tempfile,unittest,subprocess,shutil
from pathlib import Path
from PySide6.QtWidgets import QApplication
from applications import family_pdf,event_message
APP=QApplication.instance() or QApplication([])
class ApplicationTests(unittest.TestCase):
    def test_template_validation_and_literal_values(self):
        self.assertEqual(event_message('Hello {name}: {details}',{'name':'A {person}','details':'Bring food'}),'Hello A {person}: Bring food')
        for template in ('{unknown}','{name.__class__}','{name!r}','{name:20}'):
            with self.assertRaises(ValueError):event_message(template,{})
    @unittest.skipUnless(shutil.which('pdftotext'),'pdftotext required')
    def test_confirmation_contains_every_original_answer(self):
        with tempfile.TemporaryDirectory() as tmp:
            row={'Google · Full Name':'Example Family','Google · Address':'Unit 7 <West> & Road','Google · Income':'1200','Google · Dependents':'4','Google · Notes':'Please call\nAfter 5pm','Google · Missing':''}
            path=family_pdf(row,list(row),'Example Family','abc',tmp)
            self.assertTrue(path.read_bytes().startswith(b'%PDF-'))
            content=subprocess.check_output(['pdftotext',str(path),'-'],text=True)
            for expected in ['Example Family','Unit 7 <West> & Road','1200','Dependents','After 5pm','Not provided','Corrections required']:self.assertIn(expected,content)
