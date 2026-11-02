"""PPTX marker replacement, cached PDF export, and manual WhatsApp links."""
from __future__ import annotations
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from urllib.parse import quote
from pptx import Presentation

KINDS = {'volunteer':'volunteer','志愿者':'volunteer','義工':'volunteer','sukarelawan':'volunteer',
         'donor':'donor','捐赠者':'donor','捐贈者':'donor','penderma':'donor',
         'beneficiary':'beneficiary','有困难的家庭':'beneficiary','困难家庭':'beneficiary','受助家庭':'beneficiary'}


def kind(value):
    return KINDS.get(str(value or '').strip().casefold(), '')


def paragraphs(shapes):
    for shape in shapes:
        if shape.shape_type == 6:
            yield from paragraphs(shape.shapes)
        if shape.has_text_frame:
            yield from shape.text_frame.paragraphs
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    yield from cell.text_frame.paragraphs


def all_paragraphs(prs):
    for slide in prs.slides:
        yield from paragraphs(slide.shapes)
        # Some templates put markers on their layouts/master.
    for master in prs.slide_masters:
        yield from paragraphs(master.shapes)
        for layout in master.slide_layouts:
            yield from paragraphs(layout.shapes)


def markers(path):
    result = set()
    for p in all_paragraphs(Presentation(str(path))):
        result.update(re.findall(r'\{\{[^{}]+\}\}|(?<!\w)name(?!\w)',p.text,re.I))
    return sorted(result)


def replace_paragraph(paragraph, pattern, replacement):
    """Replace split-run markers without flattening unrelated formatting."""
    runs = list(paragraph.runs)
    full = ''.join(r.text for r in runs)
    hits = list(pattern.finditer(full))
    offsets=[]; pos=0
    for run in runs:
        offsets.append((pos,pos+len(run.text)));pos+=len(run.text)
    for hit in reversed(hits):
        start,end=hit.span()
        first=next(i for i,(a,b) in enumerate(offsets) if a<=start<b)
        last=next(i for i,(a,b) in enumerate(offsets) if a<end<=b)
        a=start-offsets[first][0];b=end-offsets[last][0]
        if first==last:
            runs[first].text=runs[first].text[:a]+replacement+runs[first].text[b:]
        else:
            runs[first].text=runs[first].text[:a]+replacement
            for i in range(first+1,last):runs[i].text=''
            runs[last].text=runs[last].text[b:]
    return len(hits)


def fill_template(template, output, values, name_marker='auto'):
    prs=Presentation(str(template)); count=0
    if name_marker == 'auto':
        name_pattern=re.compile(r'\{\{\s*name\s*\}\}|(?<!\w)name(?!\w)',re.I)
    else:
        name_pattern=re.compile(re.escape(name_marker))
    # Replace non-name markers first so recipient text containing markers stays literal.
    for p in all_paragraphs(prs):
        for key in ('DATE','TYPE','ID'):
            replace_paragraph(p,re.compile(r'\{\{\s*'+key+r'\s*\}\}',re.I),str(values.get(key,'')))
        count+=replace_paragraph(p,name_pattern,str(values['NAME']))
    if not count:
        raise ValueError('No matching name marker found. Select the exact text in Template Settings (name must be editable text, not an image).')
    prs.save(str(output))
    return count


def convert_pdf(pptx, output, engine='auto', soffice=''):
    pptx=Path(pptx).resolve(); output=Path(output).resolve()
    if engine=='powerpoint' or (engine=='auto' and os.name=='nt'):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize();app=None;deck=None
        try:
            app=win32com.client.DispatchEx('PowerPoint.Application')
            deck=app.Presentations.Open(str(pptx),WithWindow=False)
            deck.SaveAs(str(output),32)
        finally:
            if deck is not None: deck.Close()
            if app is not None: app.Quit()
            pythoncom.CoUninitialize()
    else:
        exe=soffice or shutil.which('libreoffice') or shutil.which('soffice')
        if not exe:
            for candidate in (r'C:\Program Files\LibreOffice\program\soffice.exe','/Applications/LibreOffice.app/Contents/MacOS/soffice'):
                if Path(candidate).exists(): exe=candidate;break
        if not exe:
            raise RuntimeError('PDF export needs Microsoft PowerPoint (Windows) or LibreOffice Impress. Install one, then retry; choose LibreOffice in Template Settings if needed.')
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp); profile=(root/'profile').as_uri(); target=root/'pdf';target.mkdir()
            result=subprocess.run([str(exe),f'-env:UserInstallation={profile}','--headless','--convert-to','pdf:impress_pdf_Export','--outdir',str(target),str(pptx)],capture_output=True,text=True,timeout=120)
            produced=target/(pptx.stem+'.pdf')
            if result.returncode or not produced.exists():
                raise RuntimeError('LibreOffice PDF conversion failed: '+(result.stderr+'\n'+result.stdout)[-1000:])
            shutil.copy2(produced,output)
    if not output.exists() or not output.read_bytes().startswith(b'%PDF-'):
        raise RuntimeError('Converter did not produce a valid PDF.')


def certificate_signature(record, template, marker, engine):
    fields={k:str(record.get(k,'')) for k in ('id','name','type','date')}
    fields.update(template=hashlib.sha256(Path(template).read_bytes()).hexdigest(),marker=marker,engine=engine)
    return hashlib.sha256(json.dumps(fields,sort_keys=True,ensure_ascii=False).encode()).hexdigest()


def generate_one(record, templates, output_dir, previous=None, engine='auto', soffice=''):
    k=kind(record.get('type'))
    if k not in ('donor','volunteer'): return {'status':'Not applicable'}
    if not str(record.get('name','')).strip(): return {'status':'Name required'}
    cfg=templates.get(k,{})
    template=Path(cfg.get('path',''))
    if not template.is_file(): return {'status':f'{k.title()} template required'}
    marker=cfg.get('marker','auto') or 'auto'
    sig=certificate_signature(record,template,marker,engine)
    previous=previous or {}
    if previous.get('signature')==sig and Path(previous.get('pdf','')).is_file(): return previous
    folder=Path(output_dir)/k;folder.mkdir(parents=True,exist_ok=True)
    label=re.sub(r'[^\w\-]+','_',str(record['name']),flags=re.U)[:55] or 'participant'
    token=hashlib.sha256(str(record['id']).encode()).hexdigest()[:16]
    dest=folder/f'{label}_{token}.pdf'
    with tempfile.TemporaryDirectory() as temp:
        pptx=Path(temp)/'certificate.pptx';pdf=Path(temp)/'certificate.pdf'
        fill_template(template,pptx,{'NAME':str(record['name']),'TYPE':k.title(),'DATE':str(record.get('date','')),'ID':str(record['id'])},marker)
        convert_pdf(pptx,pdf,engine,soffice)
        fd,staging=tempfile.mkstemp(prefix='.certificate-',suffix='.pdf',dir=folder);os.close(fd)
        try:
            shutil.copy2(pdf,staging);os.replace(staging,dest)
        finally:
            Path(staging).unlink(missing_ok=True)
    return {'status':'PDF ready','pdf':str(dest.resolve()),'signature':sig}


def whatsapp_url(phone, name, country='60'):
    digits=re.sub(r'\D','',str(phone or ''))
    if digits.startswith('00'):digits=digits[2:]
    elif digits.startswith('0'):digits=country+digits[1:]
    message=f'Hello {name or "there"}, thank you for supporting our community. Please find your certificate attached.'
    base='https://wa.me/'+digits if digits else 'https://api.whatsapp.com/send'
    return base+'?text='+quote(message)
