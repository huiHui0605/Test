# Foodbank 负责人操作指南

本系统沿用 Python + PySide6 桌面程序及本地 Excel。界面按钮均为英文，顶部只有 **Volunteer Certificates**、**Family Applications**、**Volunteer Activities** 三个标签页。

本次实现了 Google 桌面 OAuth 登录、两个回复表格的选择和同步，以及家庭确认 PDF 的 Gmail 发送流程。没有连接过你们的真实 Google 账号或真实报名表，也没有向真实申请者发送测试邮件。上线前请按最后的验收步骤，用负责人自己的测试邮箱验证。

## 1. 先分清三个系统

- 志愿者：**Tally 表单 → Tally 原生 Google Sheets 集成 → 本程序**。本程序登录 Google 不会读取 Tally 账号。
- 家庭：**Google Form → 回复 Google Sheet → 本程序**。
- 邮件：**本程序 → 已登录并授权的 Gmail / Google Workspace 邮箱**，发送真实 PDF 附件。

普通工作人员日常只需 **Connect Google**。OAuth 客户端由技术负责人一次性配置，不需要每位工作人员建立 Service Account、下载服务账号私钥或在聊天中提供密码、token。

## 2. 把志愿者 Tally 表单接到 Google Sheets

以下步骤已在 2026-11-16 核对 [Tally 官方 Google Sheets 集成说明](https://tally.so/help/google-sheets-integration)。页面按钮和政策以后可能变化，以 Tally 实际页面为准。

1. 登录组织的 Tally 账号，打开已经发布的志愿者报名表。
2. 打开 **Integrations** 标签，在 **Google Sheets** 旁点击 **Connect**。
3. 点击 **Log in to Google**，选择组织管理的 Google 账号，按 Tally 的页面授权。
4. 填写 **Connection name**。选择 **Google account**，创建新 spreadsheet，或选择现有 spreadsheet 和 sheet。
5. 如果要把已有报名导入，检查 **Export existing submissions** 是否打开，再点击 **Save changes**。官方说明支持通过此开关导出已有回复；不要假设所有旧回复已经自动补齐。核对 Tally 已有回复数量与 Sheet 实际行数，避免另行重复导入。
6. 从 **Integrations** 打开连接的 Google Sheet，检查第一行标题，以及姓名、电话、申请者选择的参加日期、Submission ID 等实际输出列。列名以你们的表单为准；若缺少必要问题，应先在表单增加并测试。
7. 提交一份明确标记为测试的报名，选择一个参加日期。确认 Sheet 增加一行，并检查日期确实来自报名者的选择。
8. 再提交一个不同参加日期的测试报名，供本程序验证逐人日期。
9. 让实际运行桌面程序的 Google 账号至少有该表格的读取权限。

**不要把 `Submitted at` / `Timestamp` 误当参加日期。** Tally 官方说明集成中的提交时间使用 UTC；这是提交时间，与参加活动的日期是两个不同字段。本程序不会用提交时间替代参加日期。

官方说明目前该集成对所有 Tally 用户免费；本程序不依赖 Zapier、Make 或直接 Tally API。若回复未送达，可查看 Tally 集成旁的同步日志及 Delivered 状态，先排查 Tally 到 Sheet 的链路。部分未完成提交不会触发该集成。

## 3. 把家庭 Google Form 接到回复 Google Sheet

负责人提供的是家庭表单编辑链接：

https://docs.google.com/forms/d/1Z8bwgdGnaaTqHR5WCY-RQu0gD9v-o25xoUGByiqW2Ok/edit

它不是回复表格链接。开发期间未读取这个私人表单的题目，也未拿到它对应的真实回复 Sheet。

1. 用有表单编辑权限的账号打开上述链接。
2. 点击 **Responses**，再点击绿色 **Google Sheets** 图标。
3. 按 Google 提示创建回复 spreadsheet，或选择已有 spreadsheet。
4. 在打开的回复 Sheet 中，确认第一行是问题标题，下面每行是一份回复。
5. 确认有申请者的邮箱字段：可以使用表单收集的邮箱，或表单中明确要求填写的电子邮箱问题。没有邮箱时，先调整表单并测试，不能用负责人的邮箱代替申请者邮箱。
6. 提交一份虚构测试申请，确认邮箱、姓名和其他答案完整出现在 Sheet 中。
7. 确保运行桌面程序的账号能打开该 Sheet。仅能编辑 Form 不代表每个工作人员都自动获得回复 Sheet 权限。
8. 保留表格链接交给负责配置的同事；在本程序中通过列表选择这个 Sheet 和实际 worksheet。

尽量给每份报名设置固定、非空、唯一的申请编号。Google Forms 不保证自动提供独立的 Submission ID。不要把非唯一 Timestamp 当作唯一 ID。若暂时没有固定编号，只能在程序中明确勾选 **Use row numbers (append-only, no unique ID)**；此模式下禁止排序、插入或删除已有回复行，否则身份可能错配。推荐由技术负责人建立一次写入且不随排序改变的唯一编号；不要使用会随行号改变的公式编号。

## 4. 技术负责人一次性 Google Cloud 配置

### 4.1 创建 OAuth Desktop App

1. 使用组织持有、可长期管理的账号进入 [Google Cloud Console](https://console.cloud.google.com/)，创建或选择组织的项目。
2. 在 **APIs & Services → Library** 启用 **Google Sheets API**、**Google Drive API**、**Gmail API**。
3. 在 **Google Auth Platform**（旧版可能显示 OAuth consent screen）配置应用名称、支持邮箱、开发者联系方式等。
4. 配置 **Audience**。符合条件的组织 Workspace 应用可选择 Internal，仅供组织内部；否则使用 External，并在测试阶段加入实际测试工作人员的账号。
5. 在 **Data Access** 配置下表的权限，并保证授权页面与实际用途一致。
6. 在 **Clients → Create client** 选择 **Desktop app**，下载客户端 JSON。不要选择 Web application 或 Service account。
7. 把 JSON 安全放在目标电脑的受控目录，例如 `C:\FoodbankConfig\oauth_client.json`，与源码 ZIP 分开保存。
8. 启动本程序，打开 **Settings → Developer Google Setup**，选择该 JSON。也可在启动前设置环境变量 `FOODBANK_GOOGLE_CLIENT` 指向文件。此步骤保存的是路径；登录令牌不写入这个文件或 settings.json。
9. 工作人员随后使用 **Connect Google** 在默认浏览器授权。回调只监听本机 `127.0.0.1` 的临时端口，使用 PKCE，不需要公网服务器。不要把回调 URL、授权码或令牌发到聊天里。

### 4.2 本程序申请的权限

| 权限 | 用途与范围 |
| --- | --- |
| `openid`、`https://www.googleapis.com/auth/userinfo.email` | 显示当前登录账号，便于核对发信身份。 |
| `https://www.googleapis.com/auth/drive.metadata.readonly` | 列出账号可访问的 Google Sheets 名称和 ID。此权限技术上可读更广的 Drive 元数据；本程序查询仅筛选 spreadsheet，不读取其他文件内容。 |
| `https://www.googleapis.com/auth/spreadsheets.readonly` | 读取用户明确选定的 worksheet、标题行及回复；不写回 Google Sheets。 |
| `https://www.googleapis.com/auth/gmail.send` | 仅在工作人员逐封确认后发送邮件。程序不申请读取 Gmail 收件箱或已发送邮件的权限，所以不确定结果必须由工作人员去 Gmail 查看。 |

Drive 的 metadata 权限属于 Google 的 restricted scope 范围；Gmail send 和 Sheets 的读取权限也有各自的敏感权限审核要求。外部应用正式发布可能需要 OAuth 验证、隐私政策、域名/应用信息证明等；若受限权限的数据被传输到或存储在服务器，可能还涉及额外安全评估。是否豁免、审核要求和费用以 Google 对具体部署的判定为准。本程序在电脑本地保存数据，不能据此保证完全免验证。

External 应用在 **Testing** 状态时，本程序这类权限组合的刷新令牌通常约 7 天到期，需要重新登录；Google 也可能因撤销授权、密码/组织策略变化等要求再次授权。不要把测试模式作为长期无人维护的正式部署方案。组织管理员可能需要在 Workspace 中允许第三方应用及 Gmail API 访问。

发信账号必须有可用的 Gmail 邮箱，或启用了 Gmail 服务且有合适许可的 Workspace 邮箱。仅有 Google 登录、但没有 Gmail 服务的账号无法完成此发信流程。发信身份为已授权账号自身，不配置 SMTP 密码，也不自动冒用其他发件地址。遵守该邮箱的发送额度。

官方参考：[Desktop OAuth](https://developers.google.com/identity/protocols/oauth2/native-app)、[OAuth 验证](https://developers.google.com/identity/protocols/oauth2/production-readiness/sensitive-scope-verification)、[Drive 权限](https://developers.google.com/workspace/drive/api/guides/api-specific-auth)、[Gmail 权限](https://developers.google.com/workspace/gmail/api/auth/scopes)、[Gmail 发送](https://developers.google.com/workspace/gmail/api/guides/sending)。上线时技术负责人应再核对 Google 控制台的实际要求。

### 4.3 凭证与资料保存

- Windows：使用当前 Windows 用户的 Credential Manager；macOS：Keychain；Linux：Secret Service / KWallet。没有受支持的安全凭证库时，程序会停止登录并提示，不会改存明文 token。
- `data/settings.json` 保存表格选择、列映射、模板路径等设置；`data/email_delivery.sqlite3` 保存不含收件人/答案的邮件尝试状态。不要删除这个状态库来“解决”重复发送提示。
- 本地 Excel、`data/family_confirmations`、感谢状 PDF 及备份含个人资料，需按组织规则限制访问和备份；OAuth 凭证安全保存不等于 Excel/PDF 已加密。
- 不要把真实 Excel、IC、PDF、截图、客户端 JSON、凭证或 data 文件夹上传到公开仓库。源码打包工具采用白名单排除这些内容。
- **Disconnect** 从此电脑的凭证库删除授权并关闭自动同步，保留已同步本地历史。若要同时撤销 Google 侧授权，进入 Google Account → Security → Third-party connections 移除本应用。断开连接不会删除已经发送的邮件。

## 5. Windows + VS Code 安装和更新

### 首次安装

安装 64 位 Python 3.12 或 3.13、VS Code 和 Python 扩展。感谢状 PPTX 转 PDF 还需要 Microsoft PowerPoint，或 LibreOffice Impress；家庭确认 PDF 不依赖这两者。

下载源码 ZIP、解压，使用 VS Code 打开含 `main.py`、`requirements.txt`、`run.bat` 的文件夹。选择 **Terminal → New Terminal → Command Prompt**，运行：

```bat
py --version
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

以后双击 **run.bat** 即可；它会检查虚拟环境并安装 requirements 中的依赖。遇到安装失败，请保留终端错误给技术负责人，不要发送真实数据或 token。Google 登录新增依赖是 `google-auth-oauthlib` 和 `keyring`，无需工作人员单独操作它们。

### 从旧版本更新

1. 关闭程序和正在打开该工作簿的 Excel。
2. 备份整个 `data`、真实 Excel 及其 `*_backups` 文件夹、自定义 PPTX 模板。Excel 可能位于程序目录以外，也要备份。
3. 把新 ZIP 解压到单独临时目录，再更新程序目录里的 `.py` 文件、`requirements.txt`、`run.bat`、`run.sh`、说明文件、`tests`、`tools` 和虚构 `examples`。
4. **保留原 data、真实 Excel、自定义 PPTX。** 新包里的示例或默认 templates 不应覆盖你们修改过的模板。保留 `.venv` 后重新安装依赖即可。不要用 `git reset --hard`、`git clean` 删除本地工作。
5. 在原程序目录执行：

```bat
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

6. 按第 4 节配置 OAuth 客户端，点击 **Connect Google**。
7. 原服务账号 JSON 不再用于本程序登录。旧 Excel 记录、隐藏 ID、历史 donor 资料和 PDF 保留；不恢复 Donor Certificates 功能。
8. 新选择器以 worksheet 的固定 ID 识别来源。旧设置使用 worksheet 名称/range 的来源迁移时，请设置新的 **Local Excel worksheet**，例如 `Google Volunteers OAuth`、`Google Families OAuth`，保留原工作表作历史，不把两种身份规则混在一个本地表中。新表后续正常增量更新。
9. 如果只是移动程序目录导致凭证槽位变化，重新 **Connect Google** 即可；不要复制系统凭证文件。`FOODBANK_DATA_DIR` 可固定 data 路径。

本项目是 Python 源码包，不是独立 Windows exe 安装程序。

## 6. 首次登录和选择两个数据源

1. 点击 **Connect Google**。浏览器应打开 Google 官方登录/授权页面；选择组织工作人员账号，核对应用名称并授予所需权限。
2. 回到程序确认 `Google: connected as ...` 显示正确账号。若取消或未完成登录，三分钟后超时，可以重试；程序界面仍可处理本地工作。
3. 打开 **Settings → Google Sheets Connection**（首次连接成功会打开），点击 **List My Google Sheets**。
4. 左侧 **Volunteer registrations (Tally → Google Sheets)**：选择 Tally 的回复 spreadsheet，再选 worksheet。右侧家庭区同样选择 Google Form 的回复 spreadsheet 和 worksheet。
5. 每个 worksheet 第一行必须是非空、互不重复的问题标题，不得使用保留的 `__fb_` 前缀。系统列出所有标题供选择，不能根据表单编辑链接猜题目。
6. 对志愿者映射 **Name column**、**Phone column**、**Attendance date column (not Timestamp)**、**Unique registration ID column**。参加日期必须手动明确选定，优先使用 Tally 的 Submission ID 作为唯一 ID。Email 可按需要映射。
7. 对家庭映射 **Name column**、**Email column**，其他字段按实际情况选择。唯一申请 ID 优先；没有时仔细阅读并明确接受 append-only 限制。
8. 给两个来源设置不同的 **Local Excel worksheet**。不要选择已放置其他资料的本地工作表。改变来源、类型或唯一 ID 方案时换一个本地工作表名称，旧表保留作历史。
9. 勾选 **Auto sync while the app is open**，默认 60 秒。点击 **Save** 后会同步；也可随时点击主界面 **Sync Google**。
10. 在相应标签页的 **Worksheet** 下拉框选择新同步的工作表。再次启动时复用保存的来源和授权；无需每次重设。程序关闭后不会继续后台同步。

同步保留原始每一列问题/答案为 `Google · 原标题`，另提供规范化的 Name/Phone/Email/Date 供操作。新 ID 追加为新行，同一 ID 更新原行；删除源回复不会删除本地历史。导入行只读，修改请在 Google Sheet 完成后同步。底部始终留约六行，可输入手动记录；空白行不写入 Excel。

没有变化的轮询不会反复保存 Excel 或增加备份。若正在编辑本地资料，同步等待保存；冲突不会覆盖外部 Excel 改动。

## 7. 日常志愿者活动与感谢状

### Volunteer Activities

1. 切换 **Volunteer Activities**，选择对应 worksheet。
2. 点击 **Event Message**，设置 Event、Time、Location、Details 和 Message template。保留已有固定活动说明。
3. 模板可用 `{name}`、`{phone}`、`{event}`、`{date}`、`{time}`、`{location}`、`{details}`，必须包含 `{date}`。全局固定日期已取消，旧配置中的固定 date 不会用于发送。
4. Google 导入的 `{date}` 来自首次数据源映射选择的参加日期列。手动 Excel 记录须先在 **Settings → Worksheet Columns → Activity participation date (not Timestamp)** 明确选列。
5. 点击该人行最右侧 **WhatsApp Share**，核对预览中的姓名、日期、活动内容；选择 **Open WhatsApp** 后在 WhatsApp 再手动发送。Cancel 不发送。电话或参加日期缺失时先补齐，程序不会替填 Timestamp。
6. 用两个不同参加日期的测试报名检查消息，确保分别显示各自日期。

### Volunteer Certificates

感谢状原流程保留：选择 **Name column**、**Attendance date column**、**Certificate number column**。PPTX 可编辑标记为 `Name`、`Date No`、`Cert No`，支持跨文字 runs 的标记。感谢状编号来自选定列，`0007` 等编号请以文本保存，不会自动编造编号。

模板需要的日期或编号缺失时不会输出不完整证书。姓名、日期、编号、模板变化会使旧缓存失效。**WhatsApp Share** 打开该人的聊天，PDF 附件和最终发送仍须人工完成。已退休 donor 数据留存，但不生成 donor 感谢状。

## 8. 家庭 PDF 预览与确认邮件

1. 切换 **Family Applications**，选择家庭工作表，找到申请者。
2. 检查该行映射的 Email；更正导入邮箱需要先改 Google Sheet 再同步，不要在邮件预览中换成别人的地址。
3. 点击最右侧 **Review & Email**。程序生成包含全部原始问题/答案的 PDF，包括空答案的 `(Not provided)`。上传附件类回答以表格中的链接/文字保留，不会额外下载私人文件。
4. 核对 **From**、**To (applicant email)**、**Subject** 和 **Message**。收件人来自该行且只读；主题和正文可以修改。
5. 点击 **Open PDF Preview** 打开真实 PDF，逐项核对。PDF 明确说明它用于资料核对，不是援助批准。
6. 勾选 **I have reviewed the recipient, message and PDF attachment.**，再点击 **Send Email**。此时才调用 Gmail API，将真正的 PDF 作为附件发送。点 **Cancel** 不发送。
7. 新报名同步不会自动发邮件。发送期间同一行按钮禁用；即使重复点击也不创建第二次投递。
8. `Email sent` 表示 Gmail API 返回了发送成功结果，不保证对方已阅读或邮件最终未被退回。
9. `Email failed` 表示已知未完成发送，先检查授权/Gmail服务/网络再重新核对。
10. `Delivery uncertain — check Gmail Sent` 表示断网、超时、服务器异常或程序上次中断导致结果不明。**先去 Gmail 的 Sent 检查，再决定是否重发。** 本程序不会自动重试。Gmail API 没有为这里提供真正的跨故障 exactly-once 保证。
11. 有任何历史发送尝试时，重发还需要勾选 **I checked Gmail Sent and explicitly want to send another email.**。不要为了消除提示删除 `email_delivery.sqlite3`。

## 9. 常见问题

| 提示或现象 | 处理 |
| --- | --- |
| 没有 OAuth Desktop App 配置 | 请技术负责人完成 Developer Google Setup；无需创建 Service Account。 |
| Google authorization expired / revoked | 点击 Connect Google，重新授权全部所需权限。检查测试模式约 7 天限制。 |
| Google denied access | 核对账号、Sheet 共享、三个 API 是否启用、Workspace 管理策略和授权范围。 |
| 表格列表没有目标 Sheet | 检查登录账号能否直接打开该 Sheet；共享后再次 List My Google Sheets。列表只显示原生 Google Sheets。 |
| worksheet 无标题或重名标题 | 修正第一行标题为非空且唯一，再选择。 |
| IDs are blank or duplicated | 修正唯一 ID；不能自动合并两份不同申请。不要改用姓名或电话作为未经验证的唯一编号。 |
| 缺参加日期或电话 | 补齐报名答案并 Sync Google；确认参加日期映射，不要使用 Timestamp。 |
| Cannot unlock credential store | 解锁系统凭证库，或由技术人员在 Linux 配置 Secret Service/KWallet；不接受明文回退。 |
| 无法发邮件 | 确认该账号有 Gmail 服务、已同意 gmail.send、未被组织策略/额度限制。 |
| 发送结果不确定 | 检查 Gmail Sent，再决定是否明确重发。 |
| Excel changed outside the app | 先保存一份手动编辑内容，关闭外部 Excel，Reload Excel 后解决冲突。不要强行覆盖。 |
| 关闭时提示后台任务未完成 | 等待当前同步/发信/证书任务完成。登录未完成会超时；强制结束可能留下邮件结果不确定状态。 |

## 10. 上线前负责人验收

请在配置好的目标 Windows 电脑使用虚构资料、自己的测试邮箱和测试电话完成：

1. 启动、登录、关闭重开，验证无需重新选择 Sheet；断开后能重新登录。
2. 新增两个不同参加日期的 Tally 回复，确认只新增两行；修改其中一个再同步，确认更新原行，无重复。
3. 新增一份家庭测试回复，核对全部字段与 PDF；取消一次发送，确认无邮件。
4. 明确发送一次到负责人自己的测试邮箱，检查正文、真实 PDF 附件、附件内容和 Gmail Sent。尝试重发时应看到额外确认。
5. 生成带 `Name / Date No / Cert No` 的感谢状，核对 `0007`、模板排版、长姓名及字体。
6. 检查 Excel 自动保存、六行空白输入、搜索、备份；确认原真实文件和自定义模板仍在。

开发测试使用模拟 Google/Gmail 接口和虚构资料。真实 OAuth 授权、你们的 Sheet 权限、Gmail 投递、Windows 凭证库与 PowerPoint COM 必须在目标环境完成上述验收后才可正式使用。
