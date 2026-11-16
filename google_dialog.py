"""Operator source picker. Only explicitly selected sheets have their rows read."""
import copy
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QCheckBox, QSpinBox, QFormLayout, QDialogButtonBox, QMessageBox, QGroupBox, QLineEdit)
from google_sync import worksheet_range, validate_source


class SourcePanel(QGroupBox):
    def __init__(self, kind, source, dialog):
        super().__init__('Volunteer registrations (Tally → Google Sheets)' if kind == 'volunteer' else 'Family applications (Google Form → Google Sheets)')
        self.kind = kind; self.dialog = dialog; self.original = copy.deepcopy(source or {})
        layout = QFormLayout(self)
        self.sheet = QComboBox(); self.worksheet = QComboBox()
        self.sheet.addItem('Choose a Google Sheet', None)
        self.worksheet.addItem('Choose a worksheet', None)
        for box in (self.sheet,self.worksheet):box.setMinimumContentsLength(12);box.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        layout.addRow('Response Google Sheet', self.sheet); layout.addRow('Worksheet', self.worksheet)
        self.fields = {}
        for key, label in [('name','Name column'), ('phone','Phone column'), ('email','Email column'),
                           ('date','Attendance date column (not Timestamp)'), ('key','Unique registration ID column')]:
            box = QComboBox(); box.addItem('Select a column', '')
            box.setMinimumContentsLength(12);box.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
            self.fields[key] = box; layout.addRow(label, box)
        self.fallback = QCheckBox('Use row numbers (append-only, no unique ID)')
        self.fallback.setChecked(bool(self.original.get('row_identity_ack')))
        layout.addRow(self.fallback)
        warning=QLabel('With row numbers, never sort, insert or delete existing response rows. Prefer Submission ID.');warning.setWordWrap(True);layout.addRow(warning)
        self.destination = QLineEdit(self.original.get('destination', 'Google Volunteers' if kind == 'volunteer' else 'Google Families'))
        layout.addRow('Local Excel worksheet', self.destination)
        self.sheet.currentIndexChanged.connect(self.choose_sheet)
        self.worksheet.currentIndexChanged.connect(self.choose_worksheet)

    def clear_fields(self):
        for box in self.fields.values():
            box.clear(); box.addItem('Select a column', '')

    def fill_sheets(self, sheets):
        self.sheet.blockSignals(True); self.sheet.clear(); self.sheet.addItem('Choose a Google Sheet', None)
        for entry in sheets:
            self.sheet.addItem(f"{entry['name']} · {entry['id'][-8:]}", entry['id'])
        self.sheet.blockSignals(False)
        # Restoring a saved selection is a read of an explicitly selected source.
        index = self.sheet.findData(self.original.get('spreadsheet'))
        if index > 0:
            self.sheet.setCurrentIndex(index)

    def choose_sheet(self):
        self.worksheet.blockSignals(True); self.worksheet.clear(); self.worksheet.addItem('Choose a worksheet', None); self.worksheet.blockSignals(False)
        self.clear_fields()
        sid = self.sheet.currentData()
        if not sid: return
        def complete(items):
            for item in items: self.worksheet.addItem(item['title'], item)
            old = self.original.get('worksheet_id')
            for index in range(1, self.worksheet.count()):
                if self.worksheet.itemData(index)['sheetId'] == old and sid == self.original.get('spreadsheet'):
                    self.worksheet.setCurrentIndex(index); break
        self.dialog.enqueue(lambda:self.dialog.account.worksheets(sid), complete)

    def choose_worksheet(self):
        self.clear_fields(); sheet = self.worksheet.currentData(); sid = self.sheet.currentData()
        if not sheet: return
        def complete(headers):
            for key, box in self.fields.items():
                for header in headers: box.addItem(header, header)
                if sid == self.original.get('spreadsheet') and sheet['sheetId'] == self.original.get('worksheet_id'):
                    box.setCurrentIndex(max(0, box.findData(self.original.get(key, ''))))
        self.dialog.enqueue(lambda:self.dialog.account.headers(sid, sheet), complete)

    def value(self):
        if not self.sheet.currentData(): return None
        sheet = self.worksheet.currentData()
        if sheet is None: raise ValueError('Choose the response worksheet for each selected Google Sheet.')
        result = dict(type=self.kind, spreadsheet=self.sheet.currentData(), worksheet_id=sheet['sheetId'],
                      range=worksheet_range(sheet), destination=self.destination.text().strip(),
                      row_identity_ack=self.fallback.isChecked())
        result.update({k:b.currentData() or '' for k,b in self.fields.items()})
        validate_source(result)
        # Source identity changes must not quietly mix two registrations in one local sheet.
        old = self.original
        identity = ('spreadsheet', 'worksheet_id', 'key', 'type')
        if old and any(old.get(k) != result.get(k) for k in identity) and old.get('destination') == result['destination']:
            raise ValueError('This is a different source or ID scheme. Choose a new local worksheet name to preserve existing history.')
        return result


class GoogleDialog(QDialog):
    def __init__(self, cfg, parent=None, account=None):
        super().__init__(parent); self.setWindowTitle('Google Sheets Connection'); self.resize(1040,680)
        self.cfg = cfg; self.account = account or parent.google_account
        self.pending = []; self.busy = False; self.panels = []
        layout = QVBoxLayout(self)
        note = QLabel('Connect Google first, then choose the two response Sheets. Only chosen sources are synchronized.\nTally registrations must reach Google Sheets through Tally Integrations. A Google Form edit link is not a response Sheet.');note.setWordWrap(True);layout.addWidget(note)
        self.refresh = QPushButton('List My Google Sheets');self.refresh.clicked.connect(self.list_sheets);layout.addWidget(self.refresh)
        columns = QHBoxLayout();layout.addLayout(columns)
        for kind in ('volunteer', 'beneficiary'):
            source = next((s for s in cfg.get('sources',[]) if s.get('type') == kind), None)
            panel = SourcePanel(kind, source, self); self.panels.append(panel);columns.addWidget(panel)
        self.enabled = QCheckBox('Auto sync while the app is open');self.enabled.setChecked(cfg.get('enabled',True));layout.addWidget(self.enabled)
        self.interval = QSpinBox();self.interval.setRange(30,3600);self.interval.setValue(cfg.get('interval',60));self.interval.setSuffix(' seconds');layout.addWidget(self.interval)
        self.status = QLabel('Click List My Google Sheets to load available response spreadsheets.');self.status.setWordWrap(True);layout.addWidget(self.status)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.validate);self.buttons.rejected.connect(self.reject);layout.addWidget(self.buttons)
        self.loaded = False

    def list_sheets(self):
        def complete(items):
            self.loaded = True
            for panel in self.panels: panel.fill_sheets(items)
            if not items: self.status.setText('No Google Sheets found for this account. Create or share the response Sheet, then refresh.')
        self.enqueue(self.account.spreadsheets, complete)

    def enqueue(self, fn, complete):
        self.pending.append((fn, complete)); self.run_next()

    def run_next(self):
        if self.busy or not self.pending: return
        self.busy = True;self.set_controls(False);self.status.setText('Reading Google selection...')
        fn, complete = self.pending.pop(0)
        def done(result):
            try: complete(result)
            finally:
                self.busy = False;self.set_controls(True)
                if self.pending: self.run_next()
                else: self.status.setText('Select fields explicitly. Attendance date means the applicant’s chosen participation date.')
        def failed(message):
            self.pending.clear();self.busy=False;self.set_controls(True);self.status.setText(message)
        self.parent().submit(fn, done, failed)

    def set_controls(self, enabled):
        self.refresh.setEnabled(enabled);self.buttons.setEnabled(enabled)
        for panel in self.panels:panel.setEnabled(enabled)

    def value(self):
        if not self.loaded: raise ValueError('Click List My Google Sheets and choose your sources before saving.')
        sources = [source for panel in self.panels if (source := panel.value())]
        if len(sources) != 2: raise ValueError('Choose both the volunteer and family response Sheets and worksheets.')
        if len({s['destination'].casefold() for s in sources}) != len(sources): raise ValueError('Use a different local worksheet name for each source.')
        return {'sources':sources, 'enabled':self.enabled.isChecked(), 'interval':self.interval.value()}

    def validate(self):
        try:
            self.value();self.accept()
        except ValueError as error:QMessageBox.warning(self,'Check Source Selection',str(error))

    def reject(self):
        if self.busy:self.status.setText('Please wait for the Google request to finish before closing.');return
        super().reject()
