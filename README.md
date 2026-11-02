# Foodbank Community Admin / 食物银行社区管理系统

PySide6 桌面系统，以 Excel 为数据库。支持志愿者、捐赠者、困难家庭资料；PPTX 感谢状自动生成 PDF；Google Sheets 单向自动同步；WhatsApp 手动分享。**不使用 Streamlit。**

## Windows 快速开始（推荐）

1. 安装 **Python 3.12 / 3.13（64 位）**，以及 **Microsoft PowerPoint**。如没有 PowerPoint，可安装 LibreOffice（需包含 Impress）。PowerPoint 更适合严格保持原 PPTX 字体和排版。
2. 解压到你有写入权限的文件夹；双击 `run.bat`。首次启动联网安装依赖。以后可直接运行 `.venv\Scripts\python.exe main.py`。
3. 系统自动创建 `data/community.xlsx`，并显示 Excel 式表格。可以直接编辑、按 Tab 换格、从 Excel 复制并用 Ctrl+V 粘贴多行多列。
4. 点击 **感谢状模板**，分别选择志愿者和捐赠者 PPTX。默认姓名标记 `auto` 会检测 `name` / `Name` / `{{NAME}}`（不区分大小写）。若模板写的是其他文字，填写精确文字即可。
5. 点击 **列映射 / 类别**，选择姓名、类别、电话和日期列。例如姓名列可以是 `Participant Name`、`姓名` 或任何自定义标题。没有 Type 列时，可给整张工作表指定默认类别。
6. 新增资料后，默认在停止输入约 1.4 秒时自动保存；姓名及类别完整后会后台生成 PDF。也会在启动、外部 Excel 变动、Google 同步后检查所有可见工作表。更改姓名、日期、类别、模板后，下次检查会重新生成；已有相同内容的 PDF 会复用。
7. 最右边 **WhatsApp ↗** 打开聊天；若该用户 PDF 已就绪，同时打开 PDF 所在文件夹并复制文件路径。**请在 WhatsApp 手动添加附件并发送**。普通 WhatsApp 链接不能自动附加本地 PDF，系统不会自动发送信息。

类别可选 `volunteer` / `donor` / `beneficiary`，也识别“志愿者”“捐赠者”“困难家庭”等。困难家庭不会生成感谢状，但可以联系。

## 先用示例体验

- `examples/sample_records.xlsx`：仅含三条虚构测试资料，不含真实电话。
- `examples/volunteer_sample.pptx`、`examples/donor_sample.pptx`：明确标记为 DEMO 的示例模板，不代表机构正式设计。
- `templates/volunteer_template.pptx`：仓库原有的志愿者模板，系统会自动发现。
- **没有提供正式捐赠者模板**。请在“感谢状模板”中选择你的正式文件，或先选择 `examples/donor_sample.pptx` 测试。
- 点击“打开 Excel”选中示例工作簿后，编辑会写入该工作簿并自动备份。若要重复体验，请先复制示例文件。

首次配置后不用每次重新选择 Excel、模板或列映射；系统会记住设置。

## Google Forms / Google Sheets 连接

1. Google 表单的报名回复先关联到 Google Sheet。
2. 在 Google Cloud 项目启用 **Google Sheets API**；创建 Service Account，并下载其 JSON key 到本机安全位置。
3. 打开 JSON，仅在本机查看 `client_email`，把需要读取的 Google Sheet 分享给这个邮箱，授予 **Viewer** 权限。不要把 JSON 上传到 Git 或发给别人。
4. 打开 **Google 连接设置**，选择 JSON，并为志愿者、困难家庭、捐赠者分别新增来源。每个来源填写：
   - 类别、Google Sheet 链接或 ID。
   - A1 范围，例如 `Form Responses 1!A1:ZZ`。第一行必须是表头；如工作表名有空格或引号，也可用 `'Form Responses 1'!A1:ZZ`。
   - 不重复的本地工作表名，例如 `Google Volunteers`、`Google Families`、`Google Donors`（最长 31 字符）。
   - 姓名、电话、Email、日期列的**精确表头**。姓名为必填，其余可留空。
   - **唯一 ID 列（建议）**：例如表单中已有的唯一报名编号。必须非空且不重复。
5. 勾选“自动同步”，默认每 60 秒轮询一次；也可以点“同步 Google”手动同步。

Google 原始列会保存在 `Google · 原标题` 列，同时生成统一的姓名、类别等列。每个来源拥有独立的本地工作表，防止混入现有资料。同步表在软件中只读；手动新增用户请在本地 Records 表操作。Google 新报名同步后会自动生成对应类别的 PDF。

同步是 **Google → 本地 Excel 单向读取**，不会修改 Google。按照唯一 ID 更新已有记录，重复轮询不重复新增。Google 中消失的行不会自动删除本地历史记录。

**未设置唯一 ID 时，以来源工作表行号为身份：Google 来源只能追加新报名，不得排序、插入或删除旧行。** 更改来源工作表、类别或身份列应使用新的本地目标工作表。日期/时间戳只有在确实唯一时才适合作为 ID。同步设置不会虚构 Google 授权；真实连接需要自己的凭证与共享权限。

Google 访问需要网络允许 `oauth2.googleapis.com` 和 `sheets.googleapis.com`。网络错误、权限错误会显示在底部，定时同步会在下次轮询重试；不影响本地 Excel 操作。

## PPTX 模板与 PDF

姓名必须是 PPTX 中**可编辑的文字**，不能只存在于背景图片中。支持文本框、组合图形、表格单元格，以及跨 run 的姓名标记。使用 `{{NAME}}` 最可靠；自动模式只匹配独立的 name 单词，避免替换 surname 等其他文字。模板上多处相同姓名标记会全部替换。可选标记：`{{DATE}}`、`{{TYPE}}`、`{{ID}}`。

原模板不会修改，系统先创建临时 PPTX，再转换为 PDF，输出到 `data/certificates/volunteer` 或 `data/certificates/donor`。姓名较长时请自行在模板中扩大文本框或缩小字号；系统保留原有字体，无法保证长姓名不会溢出。

- Windows `auto`：使用 Microsoft PowerPoint COM，需要安装 PowerPoint 和 pywin32。
- Linux / macOS `auto`：使用 LibreOffice Impress。
- Windows 没有 PowerPoint：在模板设置选择 `libreoffice`；可填写 `soffice.exe` 的完整路径。
- LibreOffice 渲染可能与 PowerPoint 有字体、换行差异；正式发送前应抽查一份 PDF 并安装模板所用字体。
- 缺模板、缺转换器、转换失败会显示具体状态。修正后点击“生成 / 重试 PDF”。后台转换不会阻塞表格操作。
- 若修改模板文件，重新载入或点击生成即可检测模板内容变化。旧姓名的历史 PDF 不自动删除，但分享按钮只使用当前资料匹配的最新 PDF。

## Excel 数据与恢复

默认文件 `data/community.xlsx`；也可“打开 Excel”直接使用现有 `.xlsx` / `.xlsm`。第一行作为表头。只适合普通表格数据；不是完整 Microsoft Excel，不能编辑图表、合并单元格布局或重新计算公式。原有未改动的数字、日期、公式保持原值；在软件中输入的文本按文本保存，不执行输入的公式。修改/删除公式引用相关行仍可能影响公式，请先检查备份。

系统为记录添加隐藏列 `__fb_id`、`__fb_source`、`__fb_key`，请保留它们，不要复制或改写 ID。每次实际写入前自动备份到 `<工作簿名>_backups`；无变化的 Google 轮询不会产生备份。备份不自动过期，管理员可自行定期归档清理。

保存使用同目录临时文件和原子替换；锁文件与时间戳检查防止大部分并发覆盖。不建议多人/多个程序同时编辑同一个 Excel。若外部文件变更且界面有未保存修改，系统停止自动保存并提示冲突；可 Ctrl+C 复制修改到其他位置后重新载入再处理。若 Excel 正在锁定文件，请关闭 Excel 后重试。

退出时会提示保存/放弃未保存的修改；后台 Google 请求或 PDF 任务完成前暂不退出。LibreOffice 转换最长等待 120 秒。

设置和 PDF 索引在 `data/settings.json`、`data/certificate_index.json`。可以用 `FOODBANK_DATA_DIR` 指定独立数据目录。请备份整个数据文件夹以及所选 Excel。丢失配置或索引可以从备份恢复；损坏 JSON 会明确报错，不会静默覆盖。

WhatsApp 默认马来西亚区号 `60`：`012...` 转为 `6012...`，`+60...` / `0060...` 支持。可在“电话区号”调整。电话号码应在 Excel 存为文本以保留开头的 0。

## Linux / macOS 安装

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
# 安装 LibreOffice Impress，例如 Debian/Ubuntu：
# sudo apt-get install libreoffice-impress
.venv/bin/python main.py
```

也可运行 `bash run.sh`。需要真实桌面显示环境；云端无桌面时可 `QT_QPA_PLATFORM=offscreen` 运行测试。桌面软件没有 Web 服务，不需要网页预览。

## 开发与测试

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m unittest discover -s tests -v
# 连同真实 PDF 导出和 pdftotext 姓名核对（需安装 poppler-utils）：
QT_QPA_PLATFORM=offscreen FOODBANK_TEST_SOFFICE=/usr/bin/libreoffice .venv/bin/python -m unittest discover -s tests -v
```

Google 自动测试使用模拟 API 响应，覆盖列映射、唯一 ID、重复同步、更新和错误输入；没有真实账号时不声称已验证你的 Google 表格。WhatsApp 测试拦截浏览器打开，不发送消息。

代码结构：`main.py` 桌面 UI/后台任务，`storage.py` Excel 数据库，`certificates.py` 模板/PDF/分享链接，`google_sync.py` Google 读取适配，`tests/` 单元及桌面测试。`tools/create_examples.py` 可重建虚构示例资料和示例模板。

源码 ZIP 不包含报名资料、服务账号凭证、缓存、虚拟环境或已生成的个人感谢状。它是可运行的 Python 源码包，并非 Windows `.exe` 安装包。
