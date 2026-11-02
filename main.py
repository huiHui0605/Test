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
    QAbstractItemView,QInputDialog,QStyledItemDelegate)
from storage import ExcelStore, ConflictError, text
from certificates import kind, markers, generate_one, certificate_signature, whatsapp_url
from google_sync import fetch

ROOT=Path(__file__).resolve().parent
DATA=Path(os.environ.get('FOODBANK_DATA_DIR',str(ROOT/'data'))).expanduser().resolve()
SETTINGS=DATA/'settings.json'
ALIASES={'name':['participantname','name','fullname','姓名','名字','namapenuh','nama'],
         'type':['type','role','category','类型','類型'],
         'phone':['phone','phonenumber','contactnumber','whatsapp','电话','telefon'],
         'date':['date','timestamp','日期','tarikh']}


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
        box=QComboBox(parent);box.setEditable(True);box.addItems(['','volunteer','donor','beneficiary']);return box
    def setEditorData(self,editor,index):editor.setCurrentText(index.data() or '')
    def setModelData(self,editor,model,index):model.setData(index,editor.currentText())


class MappingDialog(QDialog):
    def __init__(self, headers, current, parent=None):
        super().__init__(parent);self.setWindowTitle('Worksheet / 列映射');self.resize(490,280)
        layout=QFormLayout(self);self.fields={}
        for key,label in [('name','姓名 / Participant name'),('type','类别 / Type'),('phone','电话 / WhatsApp'),('date','日期 / Date')]:
            box=QComboBox();box.addItems(['(Not configured)']+headers)
            guess=next((h for h in headers if norm(h) in ALIASES[key]),'')
            chosen=current.get(key,guess)
            box.setCurrentIndex(max(0,box.findText(chosen)))
            self.fields[key]=box;layout.addRow(label,box)
        self.default=QComboBox();self.default.addItems(['Use Type column','volunteer','donor','beneficiary'])
        self.default.setCurrentText(current.get('default_type','Use Type column'))
        layout.addRow('整张表默认类别',self.default)
        label=QLabel('当表格没有 Type 列时，可为整张表选择类别。\n困难家庭不会生成感谢状。');label.setWordWrap(True);layout.addRow(label)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept);buttons.rejected.connect(self.reject);layout.addRow(buttons)
    def value(self):
        result={k:('' if b.currentIndex()==0 else b.currentText()) for k,b in self.fields.items()}
        result['default_type']=self.default.currentText();return result


class TemplateDialog(QDialog):
    def __init__(self, cfg, parent=None):
        super().__init__(parent);self.setWindowTitle('感谢状模板 / Certificate templates');self.resize(760,340)
        form=QFormLayout(self);self.paths={};self.tokens={}
        for k,label in [('volunteer','志愿者'),('donor','捐赠者')]:
            row=QWidget();line=QHBoxLayout(row);line.setContentsMargins(0,0,0,0)
            path=QLineEdit(cfg.get('templates',{}).get(k,{}).get('path',''));button=QPushButton('选择 PPTX')
            line.addWidget(path);line.addWidget(button);button.clicked.connect(lambda checked=False,key=k:self.choose(key))
            self.paths[k]=path;form.addRow(label+' PPTX',row)
            marker=QLineEdit(cfg.get('templates',{}).get(k,{}).get('marker','auto'));self.tokens[k]=marker
            form.addRow(label+'姓名标记',marker)
        self.engine=QComboBox();self.engine.addItems(['auto','powerpoint','libreoffice']);self.engine.setCurrentText(cfg.get('engine','auto'));form.addRow('PDF 转换器',self.engine)
        self.soffice=QLineEdit(cfg.get('soffice',''));form.addRow('LibreOffice 路径（可留空）',self.soffice)
        label=QLabel('auto 自动匹配 name / Name / {{NAME}}，支持跨字体段落。\n也可填写模板中的精确文字。姓名必须是可编辑文字，图片中的姓名无法替换。\nWindows auto 使用 PowerPoint；Linux/macOS 使用 LibreOffice。');label.setWordWrap(True);form.addRow(label)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.validate);buttons.rejected.connect(self.reject);form.addRow(buttons)
    def choose(self,k):
        path,_=QFileDialog.getOpenFileName(self,'选择 PPTX',str(ROOT/'templates'),'PowerPoint (*.pptx)')
        if path:
            self.paths[k].setText(path)
            try:
                found=markers(path)
                QMessageBox.information(self,'姓名标记检测','检测到：'+', '.join(found) if found else '没有检测到标准标记，请填入要替换的精确文字。')
            except Exception as e:QMessageBox.warning(self,'无法读取模板',str(e))
    def validate(self):
        for k,box in self.paths.items():
            if box.text().strip():
                try:markers(box.text().strip())
                except Exception as e:
                    QMessageBox.warning(self,'无效模板',str(e));return
        self.accept()
    def value(self):
        return {'templates':{k:{'path':box.text().strip(),'marker':self.tokens[k].text().strip() or 'auto'} for k,box in self.paths.items()},'engine':self.engine.currentText(),'soffice':self.soffice.text().strip()}


class GoogleDialog(QDialog):
    columns=[('type','类别'),('spreadsheet','Google Sheet URL / ID'),('range','范围（含表头）'),('destination','本地工作表'),('name','姓名列'),('phone','电话列'),('email','Email 列'),('date','日期列'),('key','唯一 ID 列（建议）')]
    def __init__(self,cfg,parent=None):
        super().__init__(parent);self.setWindowTitle('Google Sheets 连接设置');self.resize(1150,480)
        layout=QVBoxLayout(self)
        label=QLabel('1. 在 Google Cloud 启用 Sheets API，创建 Service Account 并下载 JSON。\n2. 把 Google 表格以 Viewer 权限分享给 JSON 中的 client_email。\n3. 每个来源单独一行；填写表头的精确名称。只读取 Google，本地修改不会回写。\n没有唯一 ID 时按行号同步：Google 来源必须只追加，不能排序、插入或删除旧行。');label.setWordWrap(True);layout.addWidget(label)
        row=QHBoxLayout();self.credentials=QLineEdit(cfg.get('credentials',''));button=QPushButton('选择凭证 JSON');button.clicked.connect(self.choose)
        row.addWidget(self.credentials);row.addWidget(button);layout.addLayout(row)
        self.table=QTableWidget(0,len(self.columns));self.table.setHorizontalHeaderLabels([label for _,label in self.columns]);self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents);layout.addWidget(self.table)
        for source in cfg.get('sources',[]):self.add(source)
        buttons=QHBoxLayout();add=QPushButton('新增来源');add.clicked.connect(lambda:self.add({}));delete=QPushButton('移除选中来源');delete.clicked.connect(self.remove)
        buttons.addWidget(add);buttons.addWidget(delete);self.enabled=QCheckBox('自动同步');self.enabled.setChecked(cfg.get('enabled',False));buttons.addWidget(self.enabled)
        self.interval=QSpinBox();self.interval.setRange(15,3600);self.interval.setValue(cfg.get('interval',60));self.interval.setSuffix(' 秒');buttons.addWidget(self.interval);layout.addLayout(buttons)
        done=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel);done.accepted.connect(self.validate);done.rejected.connect(self.reject);layout.addWidget(done)
    def choose(self):
        path,_=QFileDialog.getOpenFileName(self,'选择 Service Account JSON','','JSON (*.json)')
        if path:self.credentials.setText(path)
    def add(self,source):
        r=self.table.rowCount();self.table.insertRow(r)
        for c,(key,_) in enumerate(self.columns):
            if key=='type':
                box=QComboBox();box.addItems(['volunteer','beneficiary','donor']);box.setCurrentText(source.get(key,'volunteer'));self.table.setCellWidget(r,c,box)
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
                if not s['name'] or not s['range'] or not s['destination']:raise ValueError('每个来源必须填写范围、本地工作表和姓名列。')
                if s['destination'] in dest:raise ValueError('每个来源必须使用不同的本地工作表。')
                dest.add(s['destination'])
            if cfg['sources'] and not Path(cfg['credentials']).is_file():raise ValueError('请选择 Service Account JSON。')
            self.accept()
        except Exception as e:QMessageBox.warning(self,'设置未完成',str(e))


class Window(QMainWindow):
    def __init__(self):
        super().__init__();DATA.mkdir(parents=True,exist_ok=True)
        default={'workbook':str(DATA/'community.xlsx'),'mapping':{},'auto_pdf':True,'autosave':True,'country':'60','templates':{},'google':{}}
        self.cfg=read_json(SETTINGS,default)
        for k,v in default.items():self.cfg.setdefault(k,v)
        for k in ('volunteer','donor'):
            template=ROOT/'templates'/f'{k}_template.pptx'
            if template.exists():self.cfg['templates'].setdefault(k,{'path':str(template),'marker':'auto'})
        self.store=ExcelStore(self.cfg['workbook']);self.sheet='';self.rows=[];self.headers=[];self.visible=[];self.deleted=[]
        self.dirty=False;self.busy_pdf=False;self.busy_sync=False;self.pending_pdf=False;self.jobs=set();self.generation=0
        self.cache_path=DATA/'certificate_index.json';self.cache=read_json(self.cache_path,{})
        self.pool=QThreadPool(self);self.pool.setMaxThreadCount(2)
        self.setWindowTitle('Foodbank · Community Admin');self.resize(1380,850)
        root=QWidget();self.setCentralWidget(root);layout=QVBoxLayout(root);layout.setContentsMargins(22,18,22,16);layout.setSpacing(12)
        title=QLabel('FOODBANK  /  社区资料管理');title.setObjectName('title');layout.addWidget(title)
        self.summary=QLabel();layout.addWidget(self.summary)
        toolbar=QHBoxLayout();layout.addLayout(toolbar)
        for label,fn in [('打开 Excel',self.open_excel),('新增用户',self.add_row),('删除选中',self.delete_rows),('保存',self.save),('重新载入',self.reload),('同步 Google',self.sync_google)]:
            b=QPushButton(label);b.clicked.connect(fn);toolbar.addWidget(b)
        toolbar.addStretch();self.autosave=QCheckBox('自动保存');self.autosave.setChecked(self.cfg['autosave']);self.autosave.toggled.connect(self.options_changed);toolbar.addWidget(self.autosave)
        self.auto_pdf=QCheckBox('自动生成 PDF');self.auto_pdf.setChecked(self.cfg['auto_pdf']);self.auto_pdf.toggled.connect(self.options_changed);toolbar.addWidget(self.auto_pdf)
        filters=QHBoxLayout();layout.addLayout(filters)
        self.sheets=QComboBox();self.sheets.currentTextChanged.connect(self.change_sheet);filters.addWidget(self.sheets)
        self.search=QLineEdit();self.search.setPlaceholderText('搜索姓名、电话或任意资料…');self.search.textChanged.connect(self.render);filters.addWidget(self.search,1)
        self.category=QComboBox();self.category.addItems(['全部类别','volunteer','donor','beneficiary']);self.category.currentTextChanged.connect(self.render);filters.addWidget(self.category)
        self.table=RecordTable();self.table.setAlternatingRowColors(True);self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows);self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive);self.table.verticalHeader().setDefaultSectionSize(38);self.table.itemChanged.connect(self.edited);layout.addWidget(self.table,1)
        footer=QHBoxLayout();layout.addLayout(footer)
        for label,fn in [('列映射 / 类别',self.configure_mapping),('感谢状模板',self.configure_templates),('生成 / 重试 PDF',self.manual_pdf),('Google 连接设置',self.configure_google),('打开 PDF 文件夹',self.open_output),('电话区号',self.configure_country)]:
            b=QPushButton(label);b.clicked.connect(fn);footer.addWidget(b)
        self.info=QLabel('准备就绪');self.info.setWordWrap(True);layout.addWidget(self.info)
        self.setStyleSheet('''QMainWindow,QWidget {background:#f3f6fa;color:#172b4d;font-size:13px;} QLabel#title {font-size:24px;font-weight:700;color:#125f65;} QPushButton {background:#176b70;color:white;border:0;border-radius:5px;padding:9px 13px;} QPushButton:hover {background:#0b5157;} QPushButton:disabled {background:#95a5ab;} QLineEdit,QComboBox,QSpinBox {background:white;color:#172b4d;padding:7px;border:1px solid #cbd5e1;border-radius:4px;} QTableWidget {background:white;alternate-background-color:#edf4f7;color:#172b4d;gridline-color:#dbe3ea;selection-background-color:#c2e7e9;selection-color:#172b4d;} QHeaderView::section {background:#244753;color:white;padding:9px;border:0;} QCheckBox {spacing:6px;}''')
        self.save_timer=QTimer(self);self.save_timer.setSingleShot(True);self.save_timer.setInterval(1400);self.save_timer.timeout.connect(lambda:self.save(quiet=True))
        self.watch_timer=QTimer(self);self.watch_timer.timeout.connect(self.watch);self.watch_timer.start(2500)
        self.google_timer=QTimer(self);self.google_timer.timeout.connect(self.sync_google)
        self.load_table(self.cfg.get('sheet'))
        self.setup_google_timer()
        QTimer.singleShot(200,self.auto_generate)
        if self.cfg['google'].get('enabled'):QTimer.singleShot(500,self.sync_google)

    def persist(self):write_json(SETTINGS,self.cfg)
    def say(self,message):self.info.setText(message)
    def fail(self,message):self.say(message);QMessageBox.warning(self,'操作未完成',message)
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
        self.sheets.blockSignals(True);self.sheets.clear();self.sheets.addItems(self.store.sheets);self.sheets.setCurrentText(self.sheet);self.sheets.blockSignals(False)
        self.cfg['sheet']=self.sheet;self.cfg['workbook']=str(self.store.path);self.persist();self.render()
    def render(self,*args):
        q=self.search.text().casefold();category=self.category.currentText()
        self.visible=[i for i,r in enumerate(self.rows) if q in ' '.join(text(r.get(h)) for h in self.headers).casefold() and (category=='全部类别' or self.record(r)['type']==category)]
        self.table.blockSignals(True);self.table.clearContents();self.table.setRowCount(len(self.visible));self.table.setColumnCount(len(self.headers)+2);self.table.setHorizontalHeaderLabels(self.headers+['感谢状状态','分享'])
        for ui,i in enumerate(self.visible):
            row=self.rows[i]
            for c,h in enumerate(self.headers):
                item=QTableWidgetItem(text(row.get(h)))
                if self.sheet_is_synced():item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(ui,c,item)
            rec=self.record(row);entry=self.cache.get(self.cache_key(rec['id']),{})
            status='不需要感谢状' if rec['type']=='beneficiary' else entry.get('status','等待处理' if rec['type'] else '请选择类别 / 列映射')
            if self.dirty and rec['type'] in ('donor','volunteer'):status='待保存'
            item=QTableWidgetItem(status);item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable);self.table.setItem(ui,len(self.headers),item)
            button=QPushButton('WhatsApp ↗');button.clicked.connect(lambda checked=False,index=i:self.share(index));self.table.setCellWidget(ui,len(self.headers)+1,button)
        self.table.blockSignals(False)
        for c,h in enumerate(self.headers):
            if h==self.mapping().get('type'):self.table.setItemDelegateForColumn(c,TypeDelegate(self.table))
            else:self.table.setItemDelegateForColumn(c,None)
        self.table.setColumnWidth(0,230);self.table.setColumnWidth(len(self.headers),230);self.table.setColumnWidth(len(self.headers)+1,140)
        self.summary.setText(f'{self.store.path}  ·  {self.sheet}  ·  {len(self.rows)} 条资料 / 显示 {len(self.visible)} 条'+('  ·  未保存' if self.dirty else ''))
    def sheet_is_synced(self):return any(r.get('__fb_source') for r in self.rows) or any(s.get('destination')==self.sheet for s in self.cfg['google'].get('sources',[]))
    def edited(self,item):
        if item.column()>=len(self.headers):return
        index=self.visible[item.row()];self.rows[index][self.headers[item.column()]]=item.text();self.dirty=True;self.generation+=1
        self.say('正在编辑…'+('停止输入后自动保存。' if self.autosave.isChecked() else '请点击保存。'))
        if self.autosave.isChecked():self.save_timer.start()
    def options_changed(self):
        self.cfg['autosave']=self.autosave.isChecked();self.cfg['auto_pdf']=self.auto_pdf.isChecked();self.persist()
        if self.dirty and self.cfg['autosave']:self.save_timer.start()
        if self.cfg['auto_pdf']:QTimer.singleShot(0,self.auto_generate)
    def guard_edits(self):
        if not self.dirty:return True
        reply=QMessageBox.question(self,'未保存的修改','先保存修改吗？',QMessageBox.StandardButton.Save|QMessageBox.StandardButton.Discard|QMessageBox.StandardButton.Cancel)
        if reply==QMessageBox.StandardButton.Save:return self.save()
        if reply==QMessageBox.StandardButton.Discard:self.save_timer.stop();return True
        return False
    def save(self,checked=False,quiet=False):
        self.save_timer.stop()
        if not self.dirty:return True
        try:
            self.store.save_rows(self.sheet,self.headers,self.rows,self.deleted)
            self.load_table(self.sheet);self.say('已保存；原文件备份位于 '+str(self.store.backups))
            self.auto_generate();return True
        except Exception as e:
            message=f'保存失败，修改仍保留在界面：{e}'
            if quiet:self.say(message)
            else:self.fail(message)
            return False
    def reload(self):
        if self.guard_edits():
            self.save_timer.stop();self.store.reload();self.load_table(self.sheet);self.auto_generate()
    def open_excel(self):
        if not self.guard_edits():return
        path,_=QFileDialog.getOpenFileName(self,'打开 Excel 数据库',str(DATA),'Excel (*.xlsx *.xlsm)')
        if not path:return
        try:
            store=ExcelStore(path);self.save_timer.stop();self.store=store;self.load_table();self.auto_generate()
            self.say('已打开 Excel。保存时会自动备份，并加入隐藏的记录 ID 列。')
        except Exception as e:self.fail(str(e))
    def change_sheet(self,name):
        if not name or name==self.sheet:return
        if not self.guard_edits():
            self.sheets.blockSignals(True);self.sheets.setCurrentText(self.sheet);self.sheets.blockSignals(False);return
        self.save_timer.stop();self.load_table(name);self.auto_generate()
    def add_row(self):
        if self.sheet_is_synced():
            self.fail('Google 同步表为只读。请切换本地 Records 表新增用户。');return
        row={h:'' for h in self.headers};row.update(__fb_id='',_row=None)
        self.rows.append(row);self.dirty=True;self.search.clear();self.category.setCurrentIndex(0);self.render();self.table.scrollToBottom()
        self.table.setCurrentCell(self.table.rowCount()-1,0);self.table.editItem(self.table.item(self.table.rowCount()-1,0))
    def paste_cells(self):
        if self.sheet_is_synced():self.say('Google 同步工作表为只读。');return
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
                    self.say('筛选中只粘贴到现有行；清除筛选后可粘贴新增行。');break
                row={h:'' for h in self.headers};row.update(__fb_id='',_row=None)
                self.rows.append(row);self.visible.append(len(self.rows)-1)
            row=self.rows[self.visible[ui]]
            for c,value in enumerate(values,startcol):
                if c<len(self.headers):row[self.headers[c]]=value
        self.dirty=True;self.generation+=1;self.render()
        if self.autosave.isChecked():self.save_timer.start()
    def configure_country(self):
        value,ok=QInputDialog.getText(self,'WhatsApp 电话区号','本地号码以 0 开头时使用的国家区号（马来西亚为 60）：',text=self.cfg.get('country','60'))
        if ok and value.strip().isdigit():self.cfg['country']=value.strip();self.persist()
    def delete_rows(self):
        if self.sheet_is_synced():self.fail('请在 Google 来源管理报名记录；同步表为只读。');return
        selected={self.visible[i.row()] for i in self.table.selectedIndexes()}
        if not selected:return
        if QMessageBox.question(self,'删除资料',f'删除 {len(selected)} 条资料？保存前可以重新载入取消。')!=QMessageBox.StandardButton.Yes:return
        self.deleted.extend(self.rows[i]['_row'] for i in selected if self.rows[i].get('_row'))
        self.rows=[r for i,r in enumerate(self.rows) if i not in selected];self.dirty=True;self.render()
        if self.autosave.isChecked():self.save_timer.start()
    def configure_mapping(self):
        dialog=MappingDialog(self.headers,self.mapping(),self)
        if dialog.exec():self.cfg['mapping'][self.mapping_key()]=dialog.value();self.persist();self.generation+=1;self.render();self.auto_generate()
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
                if not any(self.record(r,mapping)['type'] in ('donor','volunteer') and self.record(r,mapping)['name'].strip() for r in rows):continue
                self.store.ensure_ids(sheet)
                _,rows=self.store.table(sheet)
                records.extend(self.record(r,mapping) for r in rows)
            self.headers,self.rows=self.store.table(self.sheet)
        except Exception as e:self.say('自动生成暂停：'+str(e));return
        records=[r for r in records if r['type'] in ('donor','volunteer') and r['name'].strip()]
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
        self.busy_pdf=True;self.say(f'后台检查 / 生成 {len(records)} 份感谢状…')
        def complete(results):
            self.busy_pdf=False;self.cache.update(results);write_json(self.cache_path,self.cache)
            if version==self.generation:self.render()
            ready=sum(v.get('status')=='PDF ready' for v in results.values())
            issues=[v['status'] for v in results.values() if v.get('status')!='PDF ready']
            self.say(f'感谢状：{ready} 份可用'+('；'+issues[0] if issues else ''))
            if self.pending_pdf:self.pending_pdf=False;QTimer.singleShot(0,self.auto_generate)
        def failed(message):self.busy_pdf=False;self.say('PDF 任务失败：'+message)
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
        if self.dirty:self.say('Google 同步等待本地修改保存。');return
        cfg=copy.deepcopy(self.cfg['google'])
        if not cfg.get('sources'):self.say('请先在 Google 连接设置中添加报名表来源。');return
        self.busy_sync=True;target=self.store.path
        def complete(batches):
            self.busy_sync=False
            if self.dirty or self.store.path!=target or self.cfg['google']!=cfg:
                self.say('Google 资料已读取；因正在编辑或切换数据库，稍后会重试同步。');return
            try:
                added,updated=self.store.merge_google(batches);current=self.sheet;self.load_table(current)
                self.say(f'Google 同步完成：新增 {added}，更新 {updated}。')
                # Process every synced worksheet, including sources not currently visible.
                if self.cfg['auto_pdf']:
                    self.auto_generate()
            except Exception as e:self.say('同步写入失败：'+str(e))
        def failed(message):self.busy_sync=False;self.say('Google 同步失败：'+message)
        self.say('正在读取 Google Sheets…');self.submit(lambda:fetch(cfg),complete,failed)
    def watch(self):
        try:
            if self.store.stamp()!=self.store.version:
                if self.dirty:self.save_timer.stop();self.say('外部 Excel 已修改。请复制保留界面中的修改，再重新载入以解决冲突。')
                else:self.store.reload();self.load_table(self.sheet);self.auto_generate()
        except Exception as e:self.say('监测 Excel 失败：'+str(e))
    def share(self,index):
        if self.dirty and not self.save():return
        if index>=len(self.rows):return
        rec=self.record(self.rows[index]);entry=self.cache.get(self.cache_key(rec['id']),{})
        current=False
        if rec['type'] in ('donor','volunteer'):
            cfg=self.cfg['templates'].get(rec['type'],{})
            try:current=entry.get('signature')==certificate_signature(rec,cfg['path'],cfg.get('marker','auto'),self.cfg.get('engine','auto')) and Path(entry.get('pdf','')).is_file()
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
        self.say('已打开 WhatsApp。请手动添加 PDF 并发送；PDF 路径已复制。' if current else '已打开 WhatsApp，请手动发送信息。感谢状尚未就绪时请检查状态并重试生成。')
    def open_output(self):
        folder=DATA/'certificates';folder.mkdir(exist_ok=True);QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))
    def closeEvent(self,event):
        if self.busy_pdf or self.busy_sync:
            self.say('后台任务仍在运行，请等待完成后关闭。');event.ignore();return
        if not self.guard_edits():event.ignore();return
        self.save_timer.stop();event.accept()


def main():
    app=QApplication(sys.argv);app.setApplicationName('Foodbank Community Admin')
    try:window=Window()
    except Exception as e:QMessageBox.critical(None,'启动失败',str(e));return 1
    window.show();return app.exec()


if __name__=='__main__':sys.exit(main())
