"""Application confirmation PDFs and configurable event messages."""
import hashlib
from html import escape
from pathlib import Path
from string import Formatter
from urllib.parse import quote
from PySide6.QtCore import QMarginsF
from PySide6.QtGui import QTextDocument,QPdfWriter,QPageSize
from certificates import whatsapp_url
from storage import text

FAMILY_FORM_URL='https://docs.google.com/forms/d/1Z8bwgdGnaaTqHR5WCY-RQu0gD9v-o25xoUGByiqW2Ok/edit'
DEFAULT_MESSAGE='Hello {name}, thank you for registering for {event}.\nDate: {date}\nTime: {time}\nLocation: {location}\n\n{details}\n\nPlease reply to confirm your attendance.'

def event_message(template,values):
    allowed={'name','event','date','time','location','details','phone'}
    for _,field,spec,conversion in Formatter().parse(template):
        if field is not None and (field not in allowed or spec or conversion):
            raise ValueError('Use only {name}, {event}, {date}, {time}, {location}, {details}, and {phone}.')
    return template.format_map({k:text(values.get(k,'')) for k in allowed})

def message_url(phone,message,country='60'):
    return whatsapp_url(phone,'',country).split('?text=')[0]+'?text='+quote(message)

def family_pdf(row,headers,name,record_id,folder):
    folder=Path(folder);folder.mkdir(parents=True,exist_ok=True)
    target=folder/(hashlib.sha256(str(record_id).encode()).hexdigest()[:24]+'.pdf')
    # Preserve the full original form response, including blank answers.
    original=[h for h in headers if h.startswith('Google · ')]
    fields=original or [h for h in headers if not h.startswith('__fb_')]
    body=''.join('<tr><td><b>'+escape(h.removeprefix('Google · '))+'</b></td><td>'+escape(text(row.get(h)) or '(Not provided)').replace('\n','<br>')+'</td></tr>' for h in fields)
    document=QTextDocument();document.setHtml('<h1>Family Application Confirmation</h1><p>Applicant: '+escape(name)+'</p><p>Please review the information below and reply on WhatsApp to confirm it is correct or describe any corrections. This document is not an approval of assistance.</p><table border="1" cellspacing="0" cellpadding="7" width="100%">'+body+'</table><p>Confirmation: Correct / Corrections required</p>')
    temporary=target.with_suffix('.tmp.pdf')
    writer=QPdfWriter(str(temporary));writer.setPageSize(QPageSize(QPageSize.PageSizeId.A4));writer.setPageMargins(QMarginsF(15,15,15,15));writer.setTitle('Family Application Confirmation');document.print_(writer);del writer
    if not temporary.read_bytes().startswith(b'%PDF-'):raise RuntimeError('Unable to create confirmation PDF.')
    temporary.replace(target);return target
