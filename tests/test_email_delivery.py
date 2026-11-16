import base64
import tempfile
import unittest
from email import policy
from email.parser import BytesParser
from pathlib import Path
from unittest.mock import Mock
from email_delivery import DeliveryJournal, DeliveryError, recipient_address, build_message, deliver, send_reviewed
from applications import activity_message, DEFAULT_MESSAGE


class MailTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
        self.pdf=self.root/'private-name.pdf';self.pdf.write_bytes(b'%PDF-1.4\nfictitious PDF test fixture')
        self.journal=DeliveryJournal(self.root/'journal.sqlite3')
    def tearDown(self):self.temp.cleanup()

    def test_mime_has_exact_recipient_and_actual_pdf_bytes(self):
        payload=build_message('family@example.test','Please confirm','Reply to confirm or correct.',self.pdf)
        message=BytesParser(policy=policy.default).parsebytes(base64.urlsafe_b64decode(payload['raw']))
        self.assertEqual(message['To'],'family@example.test')
        self.assertIsNone(message['Bcc']);self.assertIsNone(message['Cc']);self.assertIsNone(message['From'])
        attachments=list(message.iter_attachments());self.assertEqual(len(attachments),1)
        self.assertEqual(attachments[0].get_content_type(),'application/pdf')
        self.assertEqual(attachments[0].get_payload(decode=True),self.pdf.read_bytes())
        self.assertEqual(attachments[0].get_filename(),'Family-Application-Confirmation.pdf')
        self.assertIn('confirm or correct',message.get_body().get_content())

    def test_invalid_recipient_header_injection_and_non_pdf_blocked(self):
        for value in ('', 'wrong', 'a@example.test,b@example.test', 'a@example.test\nBcc: x@example.test', 'Name <a@example.test>'):
            with self.assertRaises(DeliveryError):recipient_address(value)
        with self.assertRaises(DeliveryError):build_message('a@example.test','Subject\nBcc: evil','body',self.pdf)
        self.pdf.write_bytes(b'not PDF')
        with self.assertRaises(DeliveryError):build_message('a@example.test','Subject','body',self.pdf)

    def test_journal_reserves_before_send_and_rejects_stale_preview(self):
        self.journal.reserve('record')
        with self.assertRaises(DeliveryError):self.journal.reserve('record')
        self.assertIn('uncertain',DeliveryJournal(self.journal.path).label('record'))
        self.journal.finish('record','sent','mock-message-id')
        approved=self.journal.get('record');self.journal.reserve('record',approved)
        with self.assertRaises(DeliveryError):self.journal.reserve('record',approved)
        self.assertNotIn('family@example.test',self.journal.path.read_bytes().decode('latin1'))

    def test_send_success_exactly_once_no_retries(self):
        service=Mock();send=service.users.return_value.messages.return_value.send
        send.return_value.execute.return_value={'id':'message-123'}
        result=deliver(service,{'raw':'fictional-payload'})
        self.assertEqual(result,{'status':'sent','message_id':'message-123'})
        send.assert_called_once_with(userId='me',body={'raw':'fictional-payload'})
        send.return_value.execute.assert_called_once_with(num_retries=0)

    def test_uncertain_delivery_not_retried_even_after_restart(self):
        service=Mock();service.users.return_value.messages.return_value.send.return_value.execute.side_effect=TimeoutError('secret body')
        account=Mock();account.service.return_value=service
        self.journal.reserve('record')
        result=send_reviewed(account,self.journal,'record',{'raw':'fictional-payload'})
        self.assertEqual(result['status'],'unknown')
        self.assertEqual(DeliveryJournal(self.journal.path).get('record')['status'],'unknown')
        service.users.return_value.messages.return_value.send.assert_called_once()

    def test_definitive_rejection_and_auth_failure(self):
        service=Mock();error=RuntimeError('secret recipient');error.resp=Mock(status=403)
        service.users.return_value.messages.return_value.send.return_value.execute.side_effect=error
        self.assertEqual(deliver(service,{'raw':'fictional'})['status'],'failed')
        account=Mock();account.service.side_effect=RuntimeError('credentials expired')
        self.journal.reserve('record')
        self.assertEqual(send_reviewed(account,self.journal,'record',{'raw':'fictional'})['status'],'failed')
        self.assertEqual(self.journal.get('record')['status'],'failed')

    def test_server_error_and_missing_response_id_are_uncertain(self):
        service=Mock();send=service.users.return_value.messages.return_value.send
        error=RuntimeError('private');error.resp=Mock(status=503);send.return_value.execute.side_effect=error
        self.assertEqual(deliver(service,{'raw':'fictional'})['status'],'unknown')
        send.return_value.execute.side_effect=None;send.return_value.execute.return_value={}
        self.assertEqual(deliver(service,{'raw':'fictional'})['status'],'unknown')

    def test_journal_write_failure_does_not_claim_delivery_success(self):
        service=Mock();service.users.return_value.messages.return_value.send.return_value.execute.return_value={'id':'123'}
        account=Mock();account.service.return_value=service;journal=Mock();journal.finish.side_effect=OSError('full disk')
        self.assertEqual(send_reviewed(account,journal,'record',{'raw':'fictional'})['status'],'unknown')


class ActivityTests(unittest.TestCase):
    def test_each_person_date_overrides_legacy_global_date(self):
        config={'template':DEFAULT_MESSAGE,'date':'WRONG GLOBAL DATE','event':'Packing','details':'Bring water.'}
        record={'name':'Example Volunteer','phone':'0123456789'}
        for date in ('2026-12-01','2026-12-08'):
            result=activity_message(config,record,date)
            self.assertIn(date,result);self.assertNotIn('WRONG GLOBAL DATE',result);self.assertIn('Bring water.',result)
    def test_missing_phone_date_and_date_placeholder_stop_share(self):
        config={'template':DEFAULT_MESSAGE};record={'name':'Example','phone':'0123456789'}
        for phone in ('', 'Unknown', '123', 'hello0123456789'):
            with self.assertRaises(ValueError):activity_message(config,dict(record,phone=phone),'2026-12-01')
        with self.assertRaisesRegex(ValueError,'attendance date'):activity_message(config,record,'')
        with self.assertRaisesRegex(ValueError,'must contain'):activity_message({'template':'Hello {name}'},record,'2026-12-01')


if __name__=='__main__':unittest.main()
