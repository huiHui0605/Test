"""Reviewed PDF email delivery with a durable, payload-free attempt journal."""
import base64
import re
import sqlite3
from datetime import datetime, timezone
from email.message import EmailMessage
from email.utils import format_datetime, make_msgid
from pathlib import Path


class DeliveryError(ValueError):
    pass


def recipient_address(value):
    address = str(value or '').strip()
    if len(address) > 254 or not re.fullmatch(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)+", address):
        raise DeliveryError('This applicant needs one valid email address. Correct the mapped Email field before sending.')
    return address


def build_message(recipient, subject, body, pdf):
    recipient = recipient_address(recipient)
    if not subject.strip() or '\r' in subject or '\n' in subject:
        raise DeliveryError('Enter a subject on one line.')
    if not body.strip():
        raise DeliveryError('Enter an email message.')
    data = Path(pdf).read_bytes()
    if not data.startswith(b'%PDF-'):
        raise DeliveryError('The attachment is not a PDF. Generate it again before sending.')
    if len(data) > 18 * 1024 * 1024:
        raise DeliveryError('This PDF is too large to email. Ask the administrator to review the application data.')
    message = EmailMessage()
    message['To'] = recipient
    message['Subject'] = subject
    message['Date'] = format_datetime(datetime.now(timezone.utc))
    message['Message-ID'] = make_msgid(domain='foodbank.local')
    message.set_content(body)
    message.add_attachment(data, maintype='application', subtype='pdf', filename='Family-Application-Confirmation.pdf')
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode('ascii')}


class DeliveryJournal:
    """Record status before the network call. No recipient, answers, or tokens stored."""
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connection() as db:
            db.execute('CREATE TABLE IF NOT EXISTS deliveries (record TEXT PRIMARY KEY, status TEXT NOT NULL, updated TEXT NOT NULL, message_id TEXT NOT NULL)')

    def connection(self):
        return sqlite3.connect(self.path, timeout=5)

    def get(self, key):
        with self.connection() as db:
            row = db.execute('SELECT status, updated, message_id FROM deliveries WHERE record=?', (key,)).fetchone()
        return dict(zip(('status', 'updated', 'message_id'), row)) if row else {}

    def reserve(self, key, expected=None):
        # Compare the reviewed status to current durable state, including its timestamp.
        with self.connection() as db:
            db.execute('BEGIN IMMEDIATE')
            row = db.execute('SELECT status, updated, message_id FROM deliveries WHERE record=?', (key,)).fetchone()
            current = dict(zip(('status', 'updated', 'message_id'), row)) if row else {}
            if current != (expected or {}):
                raise DeliveryError('The email status changed while you reviewed it. Close this preview, check Gmail Sent, and review again.')
            now = datetime.now(timezone.utc).isoformat()
            db.execute('INSERT OR REPLACE INTO deliveries VALUES (?, ?, ?, ?)', (key, 'sending', now, ''))

    def finish(self, key, status, message_id=''):
        with self.connection() as db:
            db.execute('UPDATE deliveries SET status=?, updated=?, message_id=? WHERE record=?', (status, datetime.now(timezone.utc).isoformat(), message_id, key))

    def label(self, key):
        entry = self.get(key)
        return {'sent':'Email sent', 'sending':'Delivery uncertain — check Gmail Sent',
                'unknown':'Delivery uncertain — check Gmail Sent', 'failed':'Email failed — review before retry'}.get(entry.get('status'), 'Ready to review & email')


def deliver(service, payload):
    """Never retry a send, including an ambiguous timeout or server error."""
    try:
        result = service.users().messages().send(userId='me', body=payload).execute(num_retries=0)
        if not result.get('id'):
            return {'status':'unknown', 'message_id':''}
        return {'status':'sent', 'message_id':result['id']}
    except Exception as error:
        status = getattr(getattr(error, 'resp', None), 'status', None)
        return {'status':'failed' if status in (400, 401, 403, 404, 413) else 'unknown', 'message_id':''}


def send_reviewed(account, journal, key, payload):
    # Credential refresh and client creation happen before Gmail's mutating call.
    try:
        service = account.service('gmail', 'v1')
    except Exception:
        result = {'status':'failed', 'message_id':''}
    else:
        try:
            result = deliver(service, payload)
        finally:
            try: service.close()
            except Exception: pass
    try:
        journal.finish(key, **result)
    except Exception:
        # The reserved 'sending' state survives; never assert success if journal update failed.
        return {'status':'unknown', 'message_id':''}
    return result
