"""Foodbank Community Admin — Excel desktop application."""
from __future__ import annotations
import copy
import hashlib
import json
import os
import sys
import webbrowser
from pathlib import Path
from PySide6.QtCore import QObject, QRunnable, QThreadPool, QTimer, Signal, Qt, QUrl
from PySide6.QtGui import QDesktopServices, QKeySequence
from PySide6.QtWidgets import (QApplication,QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,
    QLineEdit,QPushButton,QComboBox,QTableWidget,QTableWidgetItem,QHeaderView,QFileDialog,
    QMessageBox,QDialog,QFormLayout,QDialogButtonBox,QCheckBox,QSpinBox,QTabWidget,
    QAbstractItemView,QInputDialog,QStyledItemDelegate,QMenu,QStyle,QTabBar,QPlainTextEdit)
from storage import ExcelStore, ConflictError, text
from certificates import kind, markers, generate_one, certificate_signature, whatsapp_url
from google_sync import fetch
from applications import FAMILY_FORM_URL,DEFAULT_MESSAGE,event_message,message_url,family_pdf

FAMILY_TAB=1
ACTIVITY_TAB=2
ROOT=Path(__file__).resolve().parent
DATA=Path(os.environ.get('FOODBANK_DATA_DIR',str(ROOT/'data'))).expanduser().resolve()
SETTINGS=DATA/'settings.json'
ALIASES={'name':['participantname','name','fullname','姓名','名字','namapenuh','nama'],
         'type':['type','role','category','类型','類型'],
         'phone':['phone','phonenumber','contactnumber','whatsapp','电话','telefon'],
         'date':['attendancedate','dateno','date','日期','出席日期','tarikh'],
         'cert_no':['certificatenumber','certificateno','certno','感谢状编号']}


def norm(value):return ''.join(c for c in value.casefold() if c.isalnum())


def read_json(path, default):
    if not path.exists():return copy.deepcopy(default)
    try:return json.loads(path.read_text(encoding='utf-8'))
    except (ValueError,OSError) as e:raise RuntimeError(f'Cannot read {path.name}: {e}. Restore its backup instead of overwriting it.') from e


def write_json(path, value):
    path.parent.mkdir(parents=True,exist_ok=True)
    temporary=path.with_suffix('.tmp')
    temporary.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')
    os.replace(temporary,path)


class WorkerSignals(QObject):
    done=Signal(object)
    failed=Signal(str)


class Worker(QRunnable):
    def __init__(self, fn):
        super().__init__(); self.fn=fn;self.signals=WorkerSignals()
    def run(self):
        try:self.signals.done.emit(self.fn())
        except Exception as e:self.signals.failed.emit(f'{type(e).__name__}: {e}')


class RecordTable(QTableWidget):
    def keyPressEvent(self,event):
        if event.matches(QKeySequence.StandardKey.Copy):
            indexes=self.selectedIndexes()
            if indexes:
                top=min(i.row() for i in indexes);bottom=max(i.row() for i in indexes)
                left=min(i.column() for i in indexes);right=max(i.column() for i in indexes)
                selected={(i.row(),i.column()) for i in indexes}
                content='\n'.join('\t'.join(self.item(r,c).text() if (r,c) in selected and self.item(r,c) else '' for c in range(left,right+1)) for r in range(top,bottom+1))
                QApplication.clipboard().setText(content)
            return
        if event.matches(QKeySequence.StandardKey.Paste):
            owner=self.window()
            if hasattr(owner,'paste_cells'):owner.paste_cells()
            return
        super().keyPressEvent(event)


class TypeDelegate(QStyledItemDelegate):
    def createEditor(self,parent,option,index):
        box=QComboBox(parent);box.setEditable(True);box.addItems(['','volunteer','beneficiary']);return box
    def setEditorData(self,editor,index):editor.setCurrentText(index.data() or '')
    def setModelData(self,editor,model,index):model.setData(index,editor.currentText())


class MappingDialog(QDialog):
    def __init__(self, headers, current, parent=None):
        super().__init__(parent);self.setWindowTitle('Worksheet Column Mapping');self.resize(490,280)
        layout=QFormLayout(self);self.fields={}
        for key,label in [('name','Participant name'),('type','Type'),('phone','WhatsApp phone'),('date','Attendance date'),('cert_no','Certificate number')]:
            box=QComboBox();box.addItems(['(Not configured)']+headers)
            guess=next((h for h in headers if norm(h) in ALIASES[key]),'')
            chosen=current.get(key,guess)
            box.setCurrentIndex(max(0,box.findText(chosen)))
            self.fields[key]=box;layout.addRow(label,box)
        self.default=QComboBox();self.default.addItems(['Use Type column','volunteer','beneficiary'])
        self.default.setCurrentText(current.get('default_type','Use Type column'))
        layout.addRow('Default worksheet type',self.default)
        label=QLabel('Choose a default type if the worksheet has no Type column.\nBeneficiaries do not receive certificates.');label.setWordWrap(True);layout.addRow(label)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept);buttons.rejected.connect(self.reject);layout.addRow(buttons)
    def value(self):
        result={k:('' if b.currentIndex()==0 else b.currentText()) for k,b in self.fields.items()}
        result['default_type']=self.default.currentText();return result


class TemplateDialog(QDialog):
    def __init__(self, cfg, parent=None):
        super().__init__(parent);self.setWindowTitle('Certificate Templates');self.resize(760,340)
        form=QFormLayout(self);self.paths={};self.tokens={}
        for k,label in [('volunteer','Volunteer')]:
            row=QWidget();line=QHBoxLayout(row);line.setContentsMargins(0,0,0,0)
            path=QLineEdit(cfg.get('templates',{}).get(k,{}).get('path',''));button=QPushButton('Choose PPTX')
            line.addWidget(path);line.addWidget(button);button.clicked.connect(lambda checked=False,key=k:self.choose(key))
            self.paths[k]=path;form.addRow(label+' PPTX',row)
            marker=QLineEdit(cfg.get('templates',{}).get(k,{}).get('marker','Name'));self.tokens[k]=marker
            form.addRow(label+' name marker',marker)
        self.engine=QComboBox();self.engine.addItems(['auto','powerpoint','libreoffice']);self.engine.setCurrentText(cfg.get('engine','auto'));form.addRow('PDF converter',self.engine)
        self.soffice=QLineEdit(cfg.get('soffice',''));form.addRow('LibreOffice path (optional)',self.soffice)
        label=QLabel('The default name marker is Name, including text split across formatting runs. Legacy {{NAME}} templates are also supported.\nYou can also enter the exact text to replace. Use Date No for attendance date and Cert No for certificate number. All markers must be editable text.\nOn Windows, auto uses PowerPoint; on Linux/macOS, it uses LibreOffice.');label.setWordWrap(True);form.addRow(label)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.validate);buttons.rejected.connect(self.reject);form.addRow(buttons)
    def choose(self,k):
        path,_=QFileDialog.getOpenFileName(self,'Choose PPTX',str(ROOT/'templates'),'PowerPoint (*.pptx)',options=QFileDialog.Option.DontUseNativeDialog)
        if path:
            self.paths[k].setText(path)
            try:
                found=markers(path)
                QMessageBox.information(self,'Name Marker Detection','Detected: '+', '.join(found) if found else 'No standard marker found. Enter the exact text to replace.')
            except Exception as e:QMessageBox.warning(self,'Cannot Read Template',str(e))
    def validate(self):
        for k,box in self.paths.items():
            if box.text().strip():
                try:markers(box.text().strip())
                except Exception as e:
                    QMessageBox.warning(self,'Invalid Template',str(e));return
        self.accept()
    def value(self):
        return {'templates':{k:{'path':box.text().strip(),'marker':self.tokens[k].text().strip() or 'Name'} for k,box in self.paths.items()},'engine':self.engine.currentText(),'soffice':self.soffice.text().strip()}


class GoogleDialog(QDialog):
    columns=[('type','Type'),('spreadsheet','Google Sheet URL / ID'),('range','Range (including headers)'),('destination','Local worksheet'),('name','Name column'),('phone','Phone column'),('email','Email column'),('date','Date column'),('key','Unique ID column (recommended)')]
    def __init__(self,cfg,parent=None):
        super().__init__(parent);self.setWindowTitle('Google Sheets Settings');self.resize(1150,480)
        layout=QVBoxLayout(self)
        label=QLabel('Use the response Google Sheet URL, not the Google Form URL. For Tally, connect Tally to Google Sheets first.\n1. Enable the Sheets API in Google Cloud, create a service account, and download its JSON key.\n2. Share the Google Sheet with the client_email in that JSON file as a Viewer.\n3. Add one row per source and enter exact column headings. Sync is read-only; local edits are not sent to Google.\nWithout a unique ID, records use row numbers: only append responses; do not sort, insert, or delete existing rows.');label.setWordWrap(True);layout.addWidget(label)
        row=QHBoxLayout();self.credentials=QLineEdit(cfg.get('credentials',''));button=QPushButton('Choose Credentials JSON');button.clicked.connect(self.choose)
        row.addWidget(self.credentials);row.addWidget(button);layout.addLayout(row)
        self.table=QTableWidget(0,len(self.columns));self.table.setHorizontalHeaderLabels([label for _,label in self.columns]);self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents);layout.addWidget(self.table)
        for source in cfg.get('sources',[]):
            if source.get('type')!='donor':self.add(source)
        buttons=QHBoxLayout();add=QPushButton('Add Source');add.clicked.connect(lambda:self.add({}));delete=QPushButton('Remove Selected Sources');delete.clicked.connect(self.remove)
        buttons.addWidget(add);buttons.addWidget(delete);self.enabled=QCheckBox('Auto sync');self.enabled.setChecked(cfg.get('enabled',False));buttons.addWidget(self.enabled)
        self.interval=QSpinBox();self.interval.setRange(15,3600);self.interval.setValue(cfg.get('interval',60));self.interval.setSuffix(' seconds');buttons.addWidget(self.interval);layout.addLayout(buttons)
        done=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel);done.accepted.connect(self.validate);done.rejected.connect(self.reject);layout.addWidget(done)
    def choose(self):
        path,_=QFileDialog.getOpenFileName(self,'Choose Service Account JSON','','JSON (*.json)',options=QFileDialog.Option.DontUseNativeDialog)
        if path:self.credentials.setText(path)
    def add(self,source):
        r=self.table.rowCount();self.table.insertRow(r)
        for c,(key,_) in enumerate(self.columns):
            if key=='type':
                box=QComboBox();box.addItems(['volunteer','beneficiary']);box.setCurrentText(source.get(key,'volunteer'));self.table.setCellWidget(r,c,box)
            else:self.table.setItem(r,c,QTableWidgetItem(source.get(key,'Form Responses 1!A1:ZZ' if key=='range' else '')))
    def remove(self):
        for r in sorted({i.row() for i in self.table.selectedIndexes()},reverse=True):self.table.removeRow(r)
    def value(self):
        sources=[]
        for r in range(self.table.rowCount()):
            sources.append({key:(self.table.cellWidget(r,c).currentText() if key=='type' else self.table.item(r,c).text().strip()) for c,(key,_) in enumerate(self.columns)})
        return {'credentials':self.credentials.text().strip(),'sources':sources,'enabled':self.enabled.isChecked(),'interval':self.interval.value()}
    def validate(self):
        from google_sync import spreadsheet_id
        try:
            cfg=self.value();dest=set()
            for s in cfg['sources']:
                spreadsheet_id(s['spreadsheet'])
                if not s['name'] or not s['range'] or not s['destination']:raise ValueError('Each source requires a range, a local worksheet, and a name column.')
                if s['destination'] in dest:raise ValueError('Each source must use a different local worksheet.')
                dest.add(s['destination'])
            if cfg['sources'] and not Path(cfg['credentials']).is_file():raise ValueError('Please select a service account JSON file.')
            self.accept()
        except Exception as e:QMessageBox.warning(self,'Incomplete Settings',str(e))


class Window(QMainWindow):
    def __init__(self):
        super().__init__();DATA.mkdir(parents=True,exist_ok=True)
        default={'workbook':str(DATA/'community.xlsx'),'mapping':{},'auto_pdf':True,'autosave':True,'country':'60','templates':{},'google':{}}
        self.cfg=read_json(SETTINGS,default)
        for k,v in default.items():self.cfg.setdefault(k,v)
        for k in ('volunteer',):
            template=ROOT/'templates'/f'{k}_template.pptx'
            if template.exists():self.cfg['templates'].setdefault(k,{'path':str(template),'marker':'Name'})
        volunteer=self.cfg['templates'].get('volunteer')
        if volunteer is not None and volunteer.get('marker','auto') in ('','auto'):volunteer['marker']='Name'
        self.active_workflow=0;self.tab_sheets={}
        self.store=ExcelStore(self.cfg['workbook']);self.sheet='';self.rows=[];self.headers=[];self.visible=[];self.deleted=[]
        self.dirty=False;self.busy_pdf=False;self.busy_sync=False;self.pending_pdf=False;self.jobs=set();self.generation=0
        self.cache_path=DATA/'certificate_index.json';self.cache=read_json(self.cache_path,{})
        self.pool=QThreadPool(self);self.pool.setMaxThreadCount(2)
        self.setWindowTitle('Foodbank · Community Admin');self.resize(1380,850)
        root=QWidget();self.setCentralWidget(root);layout=QVBoxLayout(root);layout.setContentsMargins(22,18,22,16);layout.setSpacing(12)
        heading=QHBoxLayout();layout.addLayout(heading)
        title=QLabel('Community Records');title.setObjectName('title');heading.addWidget(title);heading.addStretch()
        settings_button=QPushButton('Settings');settings_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        settings_button.setToolTip('Set up certificates, Google Sheets, and worksheet columns.')
        self.settings_menu=QMenu(settings_button)
        for label,fn in [('Certificate Templates',self.configure_templates),('Google Sheets Connection',self.configure_google),('Worksheet Columns',self.configure_mapping),('WhatsApp Country Code',self.configure_country)]:
            self.settings_menu.addAction(label,fn)
        settings_button.setMenu(self.settings_menu);heading.addWidget(settings_button)
        self.workflow_tabs=QTabBar();self.workflow_tabs.setExpanding(True)
        for label in ['Volunteer Certificates','Family Applications','Volunteer Activities']:self.workflow_tabs.addTab(label)
        layout.addWidget(self.workflow_tabs);self.workflow_tabs.currentChanged.connect(self.change_workflow)
        guide=QLabel('Type in an empty row to add a person. Use Share to review and send on WhatsApp.')
        guide.setObjectName('guide');guide.setWordWrap(True);layout.addWidget(guide)
        self.summary=QLabel();self.summary.setWordWrap(True);layout.addWidget(self.summary)
        toolbar=QHBoxLayout();layout.addLayout(toolbar)
        for label,fn,icon,tip in [
            ('Open Excel',self.open_excel,QStyle.StandardPixmap.SP_DialogOpenButton,'Choose the Excel workbook you want to manage.'),
            ('Sync Google',self.sync_google,QStyle.StandardPixmap.SP_BrowserReload,'Check connected Google Sheets for new registrations.')]:
            button=QPushButton(label);button.setIcon(self.style().standardIcon(icon));button.setToolTip(tip);button.clicked.connect(fn);toolbar.addWidget(button)
        toolbar.addStretch()
        more=QPushButton('More Actions');more.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogListView))
        self.actions_menu=QMenu(more)
        for label,fn in [('Save Changes',self.save),('Reload Excel',self.reload),('Generate / Retry PDFs',self.manual_pdf),('Open PDF Folder',self.open_output)]:
            self.actions_menu.addAction(label,fn)
        self.actions_menu.addSeparator();self.actions_menu.addAction('Delete Selected Rows',self.delete_rows)
        more.setMenu(self.actions_menu);toolbar.addWidget(more)
        self.event_button=QPushButton('Event Message');self.event_button.clicked.connect(self.configure_event);toolbar.addWidget(self.event_button)
        self.form_button=QPushButton('Open Application Form');self.form_button.clicked.connect(lambda:webbrowser.open(FAMILY_FORM_URL));toolbar.addWidget(self.form_button)
        filters=QHBoxLayout();layout.addLayout(filters)
        filters.addWidget(QLabel('Worksheet'))
        self.sheets=QComboBox();self.sheets.setToolTip('Choose which worksheet to display.');self.sheets.currentTextChanged.connect(self.change_sheet);filters.addWidget(self.sheets)
        self.search=QLineEdit();self.search.setPlaceholderText('Search for a person...');self.search.addAction(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView),QLineEdit.ActionPosition.LeadingPosition);self.search.textChanged.connect(self.render);filters.addWidget(self.search,1)
        self.category=QComboBox();self.category.addItems(['All types','volunteer','beneficiary']);self.category.currentTextChanged.connect(self.render);filters.addWidget(self.category)
        self.certificate_columns=QWidget();column_layout=QHBoxLayout(self.certificate_columns);column_layout.setContentsMargins(0,0,0,0)
        self.column_boxes={}
        for key,label in [('name','Name column'),('date','Attendance date column'),('cert_no','Certificate number column')]:
            column_layout.addWidget(QLabel(label));box=QComboBox();box.setMinimumWidth(150);column_layout.addWidget(box,1);self.column_boxes[key]=box
            box.currentIndexChanged.connect(self.certificate_mapping_changed)
        layout.addWidget(self.certificate_columns)
        self.table=RecordTable();self.table.setAlternatingRowColors(True);self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows);self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive);self.table.verticalHeader().setDefaultSectionSize(38);self.table.itemChanged.connect(self.edited);layout.addWidget(self.table,1)
        footer=QHBoxLayout();layout.addLayout(footer)
        self.autosave=QCheckBox('Save automatically');self.autosave.setChecked(self.cfg['autosave']);self.autosave.toggled.connect(self.options_changed);footer.addWidget(self.autosave)
        self.auto_pdf=QCheckBox('Create certificates automatically');self.auto_pdf.setChecked(self.cfg['auto_pdf']);self.auto_pdf.toggled.connect(self.options_changed);footer.addWidget(self.auto_pdf)
        footer.addStretch();hint=QLabel('Double-click a cell to edit.');footer.addWidget(hint)
        self.info=QLabel('Ready');self.info.setWordWrap(True);layout.addWidget(self.info)
        self.setStyleSheet('''QMainWindow,QWidget {background:#f3f6fa;color:#172b4d;font-size:13px;} QLabel#title {font-size:26px;font-weight:700;color:#125f65;} QLabel#guide {background:#e1f0ee;color:#145b60;border-radius:8px;padding:14px;font-size:14px;} QMenu {background:white;color:#172b4d;padding:6px;border:1px solid #cbd5e1;} QMenu::item {padding:10px 24px;} QMenu::item:selected {background:#dcefed;} QPushButton {background:#176b70;color:white;border:0;border-radius:5px;padding:9px 13px;} QPushButton:hover {background:#0b5157;} QPushButton:disabled {background:#95a5ab;} QLineEdit,QComboBox,QSpinBox {background:white;color:#172b4d;padding:7px;border:1px solid #cbd5e1;border-radius:4px;} QTableWidget {background:white;alternate-background-color:#edf4f7;color:#172b4d;gridline-color:#dbe3ea;selection-background-color:#c2e7e9;selection-color:#172b4d;} QHeaderView::section {background:#244753;color:white;padding:9px;border:0;} QCheckBox {spacing:6px;}''')
        self.save_timer=QTimer(self);self.save_timer.setSingleShot(True);self.save_timer.setInterval(1400);self.save_timer.timeout.connect(lambda:self.save(quiet=True))
        self.watch_timer=QTimer(self);self.watch_timer.timeout.connect(self.watch);self.watch_timer.start(2500)
        self.google_timer=QTimer(self);self.google_timer.timeout.connect(self.sync_google)
        self.load_table(self.cfg.get('sheet'));self.update_workflow_controls()
        self.setup_google_timer()
        QTimer.singleShot(200,self.auto_generate)
        if self.cfg['google'].get('enabled'):QTimer.singleShot(500,self.sync_google)

    def workflow_kind(self):return ('volunteer','beneficiary','volunteer')[self.active_workflow]
    def update_workflow_controls(self):
        self.event_button.setVisible(self.active_workflow==ACTIVITY_TAB);self.form_button.setVisible(self.active_workflow==FAMILY_TAB)
        self.category.hide();self.certificate_columns.setVisible(self.active_workflow==0)
    def change_workflow(self,index):
        previous=self.active_workflow
        if not self.guard_edits():
            self.workflow_tabs.blockSignals(True);self.workflow_tabs.setCurrentIndex(previous);self.workflow_tabs.blockSignals(False);return
        self.save_timer.stop();self.tab_sheets[previous]=self.sheet;self.active_workflow=index
        self.search.clear();self.category.setCurrentIndex(0)
        selected=self.tab_sheets.get(index)
        if selected not in self.store.sheets:
            selected=None
            for sheet in self.store.sheets:
                headers,rows=self.store.table(sheet);mapping=self.mapping(sheet,headers)
                if any(self.record(r,mapping)['type']==self.workflow_kind() for r in rows):selected=sheet;break
        self.load_table(selected or self.sheet);self.update_workflow_controls()
    def refresh_column_boxes(self):
        mapping=self.mapping()
        for key,box in self.column_boxes.items():
            box.blockSignals(True);box.clear();box.addItem('Select a column', '');
            for header in self.headers:box.addItem(header,header)
            box.setCurrentIndex(max(0,box.findData(mapping.get(key,''))));box.blockSignals(False)
    def certificate_mapping_changed(self,*args):
        mapping=self.mapping()
        mapping.update({key:box.currentData() or '' for key,box in self.column_boxes.items()})
        self.cfg['mapping'][self.mapping_key()]=mapping;self.persist();self.generation+=1;self.render();self.auto_generate()
    def blank_row(self):
        row={h:'' for h in self.headers};row.update(__fb_id='',_row=None)
        col=self.mapping().get('type')
        if col:row[col]=self.workflow_kind()
        return row
    def append_entry_slots(self):
        start=len(self.visible);self.table.setRowCount(start+6)
        for r in range(start,start+6):
            for c in range(len(self.headers)):self.table.setItem(r,c,QTableWidgetItem(''))
            for c in (len(self.headers),len(self.headers)+1):
                item=QTableWidgetItem('');item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable);self.table.setItem(r,c,item)
    def configure_event(self):
        cfg=self.cfg.get('event_message',{});dialog=QDialog(self);dialog.setWindowTitle('Activity Details and Message');dialog.resize(650,550);form=QFormLayout(dialog);fields={}
        for key in ('event','date','time','location','details'):
            field=QLineEdit(cfg.get(key,''));fields[key]=field;form.addRow(key.title(),field)
        editor=QPlainTextEdit(cfg.get('template',DEFAULT_MESSAGE));form.addRow('Message template',editor)
        form.addRow(QLabel('Placeholders: {name}, {event}, {date}, {time}, {location}, {details}, {phone}'))
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel);buttons.accepted.connect(dialog.accept);buttons.rejected.connect(dialog.reject);form.addRow(buttons)
        if dialog.exec():
            values={k:f.text() for k,f in fields.items()};values['template']=editor.toPlainText()
            try:event_message(values['template'],values)
            except ValueError as e:self.fail(str(e));return
            self.cfg['event_message']=values;self.persist()
    def review_message(self,message,pdf=None):
        dialog=QDialog(self);dialog.setWindowTitle('Review WhatsApp Message');dialog.resize(620,480);layout=QVBoxLayout(dialog)
        layout.addWidget(QLabel('Review the message before opening WhatsApp. Send it manually there.'))
        editor=QPlainTextEdit(message);layout.addWidget(editor)
        if pdf:
            label=QLabel('Attach this PDF manually: '+str(pdf));label.setWordWrap(True);layout.addWidget(label)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel);buttons.button(QDialogButtonBox.StandardButton.Ok).setText('Open WhatsApp');buttons.accepted.connect(dialog.accept);buttons.rejected.connect(dialog.reject);layout.addWidget(buttons)
        return editor.toPlainText() if dialog.exec() else None

    def persist(self):write_json(SETTINGS,self.cfg)
    def say(self,message):self.info.setText(message)
    def fail(self,message):self.say(message);QMessageBox.warning(self,'Action Failed',message)
    def scope(self):return hashlib.sha256(str(self.store.path).encode()).hexdigest()[:20]
    def cache_key(self,uid):return self.scope()+':'+str(uid)
    def mapping_key(self, sheet=None):return self.scope()+':'+(sheet or self.sheet)
    def mapping(self, sheet=None, headers=None):
        cfg=self.cfg['mapping'].get(self.mapping_key(sheet),{})
        return {key:cfg.get(key,next((h for h in (self.headers if headers is None else headers) if norm(h) in ALIASES[key]),'')) for key in ALIASES}|{'default_type':cfg.get('default_type','Use Type column')}
    def record(self,row,mapping=None):
        mapping=mapping or self.mapping()
        rec={key:text(row.get(mapping.get(key,''))) for key in ALIASES}
        if mapping['default_type']!='Use Type column':rec['type']=mapping['default_type']
        rec['type']=kind(rec['type']);rec['id']=row.get('__fb_id','')
        return rec
    def load_table(self,sheet=None):
        self.generation+=1
        self.sheet=sheet if sheet in self.store.sheets else self.store.sheets[0]
        self.headers,self.rows=self.store.table(self.sheet);self.deleted=[];self.dirty=False
        if not self.sheet_is_synced():
            for key,label in [('date','Date'),('cert_no','Certificate Number')]:
                if not self.mapping().get(key) and label not in self.headers:self.headers.append(label)
        self.sheets.blockSignals(True);self.sheets.clear();self.sheets.addItems(self.store.sheets);self.sheets.setCurrentText(self.sheet);self.sheets.blockSignals(False)
        self.cfg['sheet']=self.sheet;self.cfg['workbook']=str(self.store.path);self.persist();self.refresh_column_boxes();self.render()
    def render(self,*args):
        q=self.search.text().casefold();category=self.category.currentText()
        self.visible=[i for i,r in enumerate(self.rows) if self.record(r)['type'] in (self.workflow_kind(),'') and q in ' '.join(text(r.get(h)) for h in self.headers).casefold() and (category=='All types' or self.record(r)['type']==category)]
        self.table.blockSignals(True);self.table.clearContents();self.table.setRowCount(len(self.visible));self.table.setColumnCount(len(self.headers)+2);self.table.setHorizontalHeaderLabels(self.headers+['Application Status' if self.active_workflow in (FAMILY_TAB,ACTIVITY_TAB) else 'Certificate Status','Share'])
        for ui,i in enumerate(self.visible):
            row=self.rows[i]
            for c,h in enumerate(self.headers):
                item=QTableWidgetItem(text(row.get(h)))
                if row.get('__fb_source'):item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(ui,c,item)
            rec=self.record(row);entry=self.cache.get(self.cache_key(rec['id']),{})
            status='No certificate required' if rec['type']=='beneficiary' else entry.get('status','Pending' if rec['type'] else 'Select a type / column mapping')
            if self.active_workflow==FAMILY_TAB:status='Ready for confirmation PDF'
            elif self.active_workflow==ACTIVITY_TAB:status='Ready to share activity details'
            if self.dirty and rec['type'] in ('volunteer',):status='Unsaved'
            item=QTableWidgetItem(status);item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable);self.table.setItem(ui,len(self.headers),item)
            button=QPushButton('Share');button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowForward));button.setToolTip('Review the activity message and open WhatsApp.' if self.active_workflow==ACTIVITY_TAB else 'Open WhatsApp. Attach the confirmation PDF and send it manually.' if self.active_workflow==FAMILY_TAB else 'Open WhatsApp. Attach the certificate PDF and send it manually.');button.clicked.connect(lambda checked=False,index=i:self.share(index));self.table.setCellWidget(ui,len(self.headers)+1,button)
        self.append_entry_slots()
        self.table.blockSignals(False)
        for c,h in enumerate(self.headers):
            if h==self.mapping().get('type'):self.table.setItemDelegateForColumn(c,TypeDelegate(self.table))
            else:self.table.setItemDelegateForColumn(c,None)
        for key,width in [('date',140),('cert_no',180)]:
            column=self.mapping().get(key)
            if column in self.headers:self.table.setColumnWidth(self.headers.index(column),width)
        self.table.setColumnWidth(0,230);self.table.setColumnWidth(len(self.headers),230);self.table.setColumnWidth(len(self.headers)+1,140)
        self.summary.setText(f'{self.store.path}  ·  {self.sheet}  ·  {len(self.rows)} records / showing {len(self.visible)}'+('  ·  Unsaved changes' if self.dirty else ''))
    def sheet_is_synced(self):return any(r.get('__fb_source') for r in self.rows) or any(s.get('destination')==self.sheet for s in self.cfg['google'].get('sources',[]))
    def edited(self,item):
        if item.column()>=len(self.headers):return
        if item.row()>=len(self.visible):
            if not item.text().strip():return
            while len(self.visible)<=item.row():
                self.rows.append(self.blank_row());self.visible.append(len(self.rows)-1)
            self.table.blockSignals(True);self.append_entry_slots();self.table.blockSignals(False)
        index=self.visible[item.row()]
        if self.rows[index].get('__fb_source'):return
        self.rows[index][self.headers[item.column()]]=item.text();self.dirty=True;self.generation+=1
        self.say('Editing... '+('Changes will save when you stop typing.' if self.autosave.isChecked() else 'Click Save to keep your changes.'))
        if self.autosave.isChecked():self.save_timer.start()
    def options_changed(self):
        self.cfg['autosave']=self.autosave.isChecked();self.cfg['auto_pdf']=self.auto_pdf.isChecked();self.persist()
        if self.dirty and self.cfg['autosave']:self.save_timer.start()
        if self.cfg['auto_pdf']:QTimer.singleShot(0,self.auto_generate)
    def guard_edits(self):
        if not self.dirty:return True
        reply=QMessageBox.question(self,'Unsaved Changes','Save your changes first?',QMessageBox.StandardButton.Save|QMessageBox.StandardButton.Discard|QMessageBox.StandardButton.Cancel)
        if reply==QMessageBox.StandardButton.Save:return self.save()
        if reply==QMessageBox.StandardButton.Discard:self.save_timer.stop();return True
        return False
    def save(self,checked=False,quiet=False):
        self.save_timer.stop()
        if not self.dirty:return True
        try:
            type_col=self.mapping().get('type')
            self.rows=[r for r in self.rows if r.get('_row') or any(text(r.get(h)).strip() for h in self.headers if h!=type_col)]
            self.store.save_rows(self.sheet,self.headers,self.rows,self.deleted)
            self.load_table(self.sheet);self.say('Saved. Workbook backups: '+str(self.store.backups))
            self.auto_generate();return True
        except Exception as e:
            message=f'Save failed; your edits remain in the table: {e}'
            if quiet:self.say(message)
            else:self.fail(message)
            return False
    def reload(self):
        if self.guard_edits():
            self.save_timer.stop();self.store.reload();self.load_table(self.sheet);self.auto_generate()
    def open_excel(self):
        if not self.guard_edits():return
        path,_=QFileDialog.getOpenFileName(self,'Open Excel Database',str(DATA),'Excel (*.xlsx *.xlsm)',options=QFileDialog.Option.DontUseNativeDialog)
        if not path:return
        try:
            store=ExcelStore(path);self.save_timer.stop();self.store=store;self.load_table();self.auto_generate()
            self.say('Excel opened. Saving creates a backup and adds hidden record ID columns.')
        except Exception as e:self.fail(str(e))
    def change_sheet(self,name):
        if not name or name==self.sheet:return
        if not self.guard_edits():
            self.sheets.blockSignals(True);self.sheets.setCurrentText(self.sheet);self.sheets.blockSignals(False);return
        self.save_timer.stop();self.load_table(name);self.auto_generate()
    def add_person_dialog(self):
        if self.sheet_is_synced():
            self.fail('Google worksheets are read-only. Choose a local worksheet to add a person.');return
        if not self.guard_edits():return
        # Discard must reset the table before the form adds a new record.
        if self.dirty:self.store.reload();self.load_table(self.sheet)
        dialog=QDialog(self);dialog.setWindowTitle('Add Person');dialog.resize(500,350)
        form=QFormLayout(dialog);fields={};mapping=self.mapping()
        note=QLabel('Enter the person details below. Certificates are created for volunteers.');note.setWordWrap(True);form.addRow(note)
        for header in self.headers:
            if header==mapping.get('type'):
                editor=QComboBox();editor.addItems(['volunteer','beneficiary'])
            else:
                editor=QLineEdit()
                if header==mapping.get('name'):editor.setPlaceholderText('Name to print on the certificate')
            editor.setObjectName('person_'+header);fields[header]=editor;form.addRow(header,editor)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept);buttons.rejected.connect(dialog.reject);form.addRow(buttons)
        if dialog.exec()!=QDialog.DialogCode.Accepted:return
        row={h:(field.currentText() if isinstance(field,QComboBox) else field.text().strip()) for h,field in fields.items()}
        if not any(row.values()):return
        row.update(__fb_id='',_row=None);self.rows.append(row);self.dirty=True
        self.search.clear();self.category.setCurrentIndex(0);self.render();self.save()

    def add_row(self):
        if self.sheet_is_synced():
            self.fail('Google worksheets are read-only. Switch to a local worksheet to add users.');return
        row={h:'' for h in self.headers};row.update(__fb_id='',_row=None)
        self.rows.append(row);self.dirty=True;self.search.clear();self.category.setCurrentIndex(0);self.render();self.table.scrollToBottom()
        self.table.setCurrentCell(self.table.rowCount()-1,0);self.table.editItem(self.table.item(self.table.rowCount()-1,0))
    def paste_cells(self):
        import csv,io
        try:matrix=list(csv.reader(io.StringIO(QApplication.clipboard().text()),delimiter='\t'))
        except csv.Error as e:self.say(str(e));return
        if not matrix:return
        startrow=max(0,self.table.currentRow());startcol=max(0,self.table.currentColumn())
        if startcol>=len(self.headers):return
        for offset,values in enumerate(matrix):
            ui=startrow+offset
            if ui>=len(self.visible):
                if self.search.text() or self.category.currentIndex()!=0:
                    self.say('While filtered, paste only fills existing rows. Clear filters to paste new rows.');break
                row=self.blank_row()
                self.rows.append(row);self.visible.append(len(self.rows)-1)
            row=self.rows[self.visible[ui]]
            if row.get('__fb_source'):
                self.say('Google response rows are read-only. Paste into the empty rows below.');continue
            for c,value in enumerate(values,startcol):
                if c<len(self.headers):row[self.headers[c]]=value
        self.dirty=True;self.generation+=1;self.render()
        if self.autosave.isChecked():self.save_timer.start()
    def configure_country(self):
        value,ok=QInputDialog.getText(self,'WhatsApp Country Code','Country code for local numbers starting with 0 (Malaysia: 60):',text=self.cfg.get('country','60'))
        if ok and value.strip().isdigit():self.cfg['country']=value.strip();self.persist()
    def delete_rows(self):
        selected={self.visible[i.row()] for i in self.table.selectedIndexes() if i.row()<len(self.visible)}
        if any(self.rows[i].get('__fb_source') for i in selected):self.fail('Google response rows cannot be deleted here. Select manual rows only.');return
        if not selected:return
        if QMessageBox.question(self,'Delete Records',f'Delete {len(selected)} records? Reload before saving to undo this change.')!=QMessageBox.StandardButton.Yes:return
        self.deleted.extend(self.rows[i]['_row'] for i in selected if self.rows[i].get('_row'))
        self.rows=[r for i,r in enumerate(self.rows) if i not in selected];self.dirty=True;self.render()
        if self.autosave.isChecked():self.save_timer.start()
    def configure_mapping(self):
        dialog=MappingDialog(self.headers,self.mapping(),self)
        if dialog.exec():self.cfg['mapping'][self.mapping_key()]=dialog.value();self.persist();self.generation+=1;self.refresh_column_boxes();self.render();self.auto_generate()
    def configure_templates(self):
        dialog=TemplateDialog(self.cfg,self)
        if dialog.exec():self.cfg.update(dialog.value());self.persist();self.generation+=1;self.auto_generate()
    def manual_pdf(self):
        if self.dirty and not self.save():return
        self.auto_generate(force=True)
    def auto_generate(self,force=False):
        if self.dirty or (not force and not self.cfg['auto_pdf']):return
        if self.busy_pdf:self.pending_pdf=True;return
        try:
            records=[]
            for sheet in self.store.sheets:
                headers,rows=self.store.table(sheet)
                mapping=self.mapping(sheet,headers)
                if not any(self.record(r,mapping)['type'] in ('volunteer',) and self.record(r,mapping)['name'].strip() for r in rows):continue
                self.store.ensure_ids(sheet)
                _,rows=self.store.table(sheet)
                records.extend(self.record(r,mapping) for r in rows)
            _,self.rows=self.store.table(self.sheet)
        except Exception as e:self.say('Automatic generation paused: '+str(e));return
        records=[r for r in records if r['type'] in ('volunteer',) and r['name'].strip()]
        if not records:self.render();return
        templates=copy.deepcopy(self.cfg['templates']);engine=self.cfg.get('engine','auto');soffice=self.cfg.get('soffice','')
        old=copy.deepcopy(self.cache);scope=self.scope();version=self.generation
        def work():
            results={}
            for rec in records:
                key=scope+':'+rec['id']
                try:results[key]=generate_one(rec,templates,DATA/'certificates',old.get(key),engine,soffice)
                except Exception as e:results[key]={'status':str(e)}
            return results
        self.busy_pdf=True;self.say(f'Checking / generating {len(records)} certificates in the background...')
        def complete(results):
            self.busy_pdf=False;self.cache.update(results);write_json(self.cache_path,self.cache)
            if version==self.generation:self.render()
            ready=sum(v.get('status')=='PDF ready' for v in results.values())
            issues=[v['status'] for v in results.values() if v.get('status')!='PDF ready']
            self.say(f'Certificates: {ready} ready'+('; '+issues[0] if issues else ''))
            if self.pending_pdf:self.pending_pdf=False;QTimer.singleShot(0,self.auto_generate)
        def failed(message):self.busy_pdf=False;self.say('PDF task failed: '+message)
        self.submit(work,complete,failed)
    def submit(self,fn,complete,failed):
        job=Worker(fn);self.jobs.add(job)
        def success(result):
            try:complete(result)
            finally:self.jobs.discard(job)
        def error(message):
            try:failed(message)
            finally:self.jobs.discard(job)
        job.signals.done.connect(success);job.signals.failed.connect(error);self.pool.start(job)
    def configure_google(self):
        dialog=GoogleDialog(self.cfg['google'],self)
        if dialog.exec():self.cfg['google']=dialog.value();self.persist();self.setup_google_timer();self.sync_google()
    def setup_google_timer(self):
        self.google_timer.stop()
        if self.cfg['google'].get('enabled'):self.google_timer.start(max(15,self.cfg['google'].get('interval',60))*1000)
    def sync_google(self):
        if self.busy_sync:return
        if self.dirty:self.say('Google sync is waiting for local edits to be saved.');return
        original_cfg=copy.deepcopy(self.cfg['google']);cfg=copy.deepcopy(original_cfg)
        cfg['sources']=[s for s in cfg.get('sources',[]) if s.get('type')!='donor']
        if not cfg.get('sources'):self.say('Add registration sources in Google Settings first.');return
        self.busy_sync=True;target=self.store.path
        def complete(batches):
            self.busy_sync=False
            if self.dirty or self.store.path!=target or self.cfg['google']!=original_cfg:
                self.say('Google data was read. Sync was deferred because edits or connection settings changed; retry when ready.');return
            try:
                added,updated=self.store.merge_google(batches);current=self.sheet;self.load_table(current)
                self.say(f'Google sync complete: {added} added, {updated} updated.')
                # Process every synced worksheet, including sources not currently visible.
                if self.cfg['auto_pdf']:
                    self.auto_generate()
            except Exception as e:self.say('Could not save synced data: '+str(e))
        def failed(message):self.busy_sync=False;self.say('Google sync failed: '+message)
        self.say('Reading Google Sheets...');self.submit(lambda:fetch(cfg),complete,failed)
    def watch(self):
        try:
            if self.store.stamp()!=self.store.version:
                if self.dirty:self.save_timer.stop();self.say('Excel changed outside the app. Copy your unsaved edits, then reload to resolve the conflict.')
                else:self.store.reload();self.load_table(self.sheet);self.auto_generate()
        except Exception as e:self.say('Could not monitor Excel: '+str(e))
    def share(self,index):
        original=self.rows[index] if index<len(self.rows) else None
        if original is None:return
        if self.dirty:
            col=self.mapping().get('type')
            new_index=sum(bool(r.get('_row') or any(text(r.get(h)).strip() for h in self.headers if h!=col)) for r in self.rows[:index])
            if not self.save():return
            index=new_index
        if index>=len(self.rows):return
        if self.active_workflow in (FAMILY_TAB,ACTIVITY_TAB):
            self.share_application(index);return
        rec=self.record(self.rows[index]);entry=self.cache.get(self.cache_key(rec['id']),{})
        current=False
        if rec['type'] in ('volunteer',):
            cfg=self.cfg['templates'].get(rec['type'],{})
            try:current=entry.get('signature')==certificate_signature(rec,cfg['path'],cfg.get('marker','Name'),self.cfg.get('engine','auto')) and Path(entry.get('pdf','')).is_file()
            except (KeyError,OSError):pass
            if current:
                QApplication.clipboard().setText(entry['pdf'])
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(entry['pdf']).parent)))
        # Explicit user click only; never send messages or attachments automatically.
        url=whatsapp_url(text(self.rows[index].get(self.mapping().get('phone',''))),rec['name'],self.cfg.get('country','60'))
        if not current:
            from urllib.parse import quote
            url=url.split('?text=')[0]+'?text='+quote(f'Hello {rec["name"] or "there"}, we are contacting you from the community foodbank.')
        webbrowser.open(url)
        self.say('WhatsApp opened. Attach the PDF and send manually; its path has been copied.' if current else 'WhatsApp opened. Send your message manually. If the certificate is not ready, check its status and retry generation.')
    def share_application(self,index):
        row=self.rows[index];rec=self.record(row);phone=rec['phone'];pdf=None
        try:
            if self.active_workflow==FAMILY_TAB:
                self.store.ensure_ids(self.sheet);self.headers,self.rows=self.store.table(self.sheet);row=self.rows[index];rec=self.record(row)
                pdf=family_pdf(row,self.headers,rec['name'],self.scope()+':'+rec['id'],DATA/'family_confirmations')
                message=f"Hello {rec['name'] or 'there'}, please review your family application details in the attached PDF. Reply to confirm they are correct or tell us what needs changing."
            else:
                cfg=self.cfg.get('event_message',{})
                if not cfg.get('event'):
                    self.configure_event();cfg=self.cfg.get('event_message',{})
                    if not cfg.get('event'):return
                message=event_message(cfg.get('template',DEFAULT_MESSAGE),dict(cfg,name=rec['name'],phone=phone))
            approved=self.review_message(message,pdf)
            if approved is None:return
            if pdf:
                QApplication.clipboard().setText(str(pdf));QDesktopServices.openUrl(QUrl.fromLocalFile(str(pdf.parent)))
            webbrowser.open(message_url(phone,approved,self.cfg.get('country','60')))
            self.say('WhatsApp opened. Attach the confirmation PDF manually.' if pdf else 'WhatsApp opened. Review and send the activity message manually.')
        except Exception as e:self.fail(str(e))

    def open_output(self):
        folder=DATA/'certificates';folder.mkdir(exist_ok=True);QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))
    def closeEvent(self,event):
        if self.busy_pdf or self.busy_sync:
            self.say('Background tasks are still running. Wait for them to finish before closing.');event.ignore();return
        if not self.guard_edits():event.ignore();return
        self.save_timer.stop();event.accept()


def main():
    app=QApplication(sys.argv);app.setApplicationName('Foodbank Community Admin')
    try:window=Window()
    except Exception as e:QMessageBox.critical(None,'Startup Failed',str(e));return 1
    window.show();return app.exec()


if __name__=='__main__':sys.exit(main())
