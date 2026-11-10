"""Excel database with atomic writes, conflict detection and backups."""
from __future__ import annotations
import os
import shutil
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path
from openpyxl import Workbook, load_workbook

META = ('__fb_id', '__fb_source', '__fb_key')
DEFAULT_HEADERS = ['Participant Name', 'Type', 'Phone', 'Email', 'Date', 'Certificate Number']


def text(value):
    if value is None:
        return ''
    if hasattr(value, 'isoformat'):
        return value.isoformat()
    return str(value)


class ConflictError(RuntimeError):
    pass


class ExcelStore:
    def __init__(self, path):
        self.path = Path(path).expanduser().resolve()
        if self.path.suffix.lower() not in ('.xlsx', '.xlsm'):
            raise ValueError('Please select an .xlsx or .xlsm workbook.')
        self.backups = self.path.parent / (self.path.stem + '_backups')
        if not self.path.exists():
            self.path.parent.mkdir(parents=True, exist_ok=True)
            w = Workbook(); w.active.title = 'Records'
            w.active.append(DEFAULT_HEADERS)
            w.save(self.path); w.close()
        self.reload()

    def stamp(self):
        s = self.path.stat()
        return (s.st_mtime_ns, s.st_size)

    def reload(self):
        self.wb = load_workbook(self.path, keep_vba=self.path.suffix.lower() == '.xlsm')
        self.version = self.stamp()
        if not self.wb.sheetnames:
            raise ValueError('Workbook has no worksheets.')

    @property
    def sheets(self):
        return [s.title for s in self.wb if s.sheet_state == 'visible']

    @staticmethod
    def display_headers(ws):
        raw = [text(c.value) for c in ws[1]]
        # Keep duplicate/blank column labels independently addressable in the UI.
        used = set(); headers = []
        for i, label in enumerate(raw, 1):
            label = label or f'Column {i}'
            base = label; n = 2
            while label in used:
                label = f'{base} ({n})'; n += 1
            headers.append(label); used.add(label)
        return headers

    def table(self, sheet):
        ws = self.wb[sheet]
        headers = self.display_headers(ws)
        indexes = [i for i,h in enumerate(headers) if h not in META]
        rows = []
        for excel_row in range(2, ws.max_row + 1):
            cells = [ws.cell(excel_row, i+1).value for i in range(len(headers))]
            if not any(cells[i] is not None for i in indexes):
                continue
            record = {h: cells[i] for i,h in enumerate(headers)}
            record['_row'] = excel_row
            rows.append(record)
        return [headers[i] for i in indexes], rows

    def commit(self, operation):
        lock = self.path.with_name(self.path.name + '.foodbank.lock')
        try:
            fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        except FileExistsError:
            raise ConflictError('Another app is saving this workbook. If an app crashed, close all instances before removing its .foodbank.lock file.')
        os.close(fd)
        tmp = None
        try:
            if self.stamp() != self.version:
                raise ConflictError('Excel was changed outside this app. Reload before saving; your edits have not been overwritten.')
            changed = operation(self.wb)
            if changed is False:
                return
            self.backups.mkdir(exist_ok=True)
            stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
            shutil.copy2(self.path, self.backups / f'{self.path.stem}_{stamp}{self.path.suffix}')
            fd,tmp = tempfile.mkstemp(prefix='.foodbank-', suffix=self.path.suffix, dir=self.path.parent)
            os.close(fd)
            self.wb.save(tmp)
            if self.stamp() != self.version:
                raise ConflictError('Workbook changed during save. Reload and resolve the changes.')
            os.replace(tmp, self.path); tmp = None
            self.reload()
        except Exception:
            self.reload()
            raise
        finally:
            if tmp and os.path.exists(tmp): os.unlink(tmp)
            lock.unlink(missing_ok=True)

    @staticmethod
    def metadata(ws):
        headers = [text(c.value) for c in ws[1]]
        for key in META:
            if key not in headers:
                headers.append(key); ws.cell(1,len(headers),key)
            from openpyxl.utils import get_column_letter
            ws.column_dimensions[get_column_letter(headers.index(key)+1)].hidden = True
        return headers

    def save_rows(self, sheet, headers, rows, deleted=()):
        """Only write changed cells; preserve original numeric/date/formula cells."""
        def mutate(wb):
            ws = wb[sheet]
            display = self.display_headers(ws)
            for header in headers:
                if header not in display:
                    if header in META:raise ValueError('Reserved metadata column')
                    ws.cell(1,ws.max_column+1,header);display=self.display_headers(ws)
            all_headers = self.metadata(ws)
            display = self.display_headers(ws)
            columns = {h:display.index(h)+1 for h in headers}
            idcol = all_headers.index('__fb_id') + 1
            seen_ids=set()
            for rec in rows:
                r = rec.get('_row') or ws.max_row + 1
                for h,c in columns.items():
                    new = rec.get(h)
                    old = ws.cell(r,c).value
                    if text(new) != text(old):
                        cell = ws.cell(r,c)
                        cell.value = new if new != '' else None
                        # Input/synced strings must never become Excel formulas.
                        if isinstance(new,str): cell.data_type = 's'
                if not ws.cell(r,idcol).value or ws.cell(r,idcol).value in seen_ids:
                    ws.cell(r,idcol,str(uuid.uuid4()))
                seen_ids.add(ws.cell(r,idcol).value)
            for r in sorted(set(deleted), reverse=True): ws.delete_rows(r)
        self.commit(mutate)

    def ensure_ids(self, sheet):
        headers, rows = self.table(sheet)
        ids=[r.get('__fb_id') for r in rows]
        if all(ids) and len(set(ids))==len(ids): return
        self.save_rows(sheet, headers, rows)

    def merge_google(self, batches):
        """Each source owns a dedicated sheet. Stable IDs preferred; append-only rows otherwise."""
        counts = [0,0]
        def mutate(wb):
            changed=False
            for batch in batches:
                title = batch['sheet']
                if title in wb.sheetnames:
                    ws = wb[title]
                    # Never take over a manually created worksheet.
                    existing = [text(c.value) for c in ws[1]]
                    if not all(k in existing for k in META):
                        raise ValueError(f'Worksheet {title!r} already exists and is not a synced worksheet. Choose a new destination name.')
                    sources = {text(ws.cell(r,existing.index('__fb_source')+1).value) for r in range(2,ws.max_row+1)} - {''}
                    if sources - {batch['source']}:
                        raise ValueError(f'Worksheet {title!r} belongs to another Google source.')
                else:
                    changed=True
                    ws = wb.create_sheet(title)
                    ws.append(batch['headers'] + list(META))
                headers = self.metadata(ws)
                for h in batch['headers']:
                    if h not in headers:
                        changed=True
                        headers.append(h); ws.cell(1,len(headers),h)
                cols = {h:i+1 for i,h in enumerate(headers)}
                lookup = {text(ws.cell(r,cols['__fb_key']).value):r for r in range(2,ws.max_row+1)}
                for rec in batch['rows']:
                    key = rec['__fb_key']; r = lookup.get(key)
                    if r is None:
                        changed=True
                        r = ws.max_row+1; lookup[key]=r; counts[0]+=1
                        ws.cell(r,cols['__fb_id'],str(uuid.uuid4()))
                    elif any(text(ws.cell(r,cols[h]).value) != text(rec.get(h)) for h in batch['headers']):
                        changed=True
                        counts[1]+=1
                    for h,v in rec.items():
                        if h == '__fb_id': continue
                        cell=ws.cell(r,cols[h]);cell.value=v
                        if isinstance(v,str):cell.data_type='s'
            return changed
        if batches: self.commit(mutate)
        return tuple(counts)
