# Foodbank Community Admin

An English-language PySide6 desktop application using Excel as its database. Manage volunteers and beneficiary families; generate PDF certificates from PPTX templates; sync Google Sheets; and open WhatsApp for manual sharing. No Streamlit or web server is required.

![Desktop interface](examples/interface.png)

## Windows: first run in VS Code

Install Python 3.12 or 3.13 (64-bit), VS Code with Microsoft's Python extension, and Microsoft PowerPoint or LibreOffice Impress for PDF export. PowerPoint generally preserves PowerPoint template formatting most accurately.

Download or clone this repository and open its root folder in VS Code. Open **Terminal → New Terminal**, select **Command Prompt**, and run:

```bat
py --version
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

For subsequent launches, run `.venv\Scripts\python.exe main.py` or double-click `run.bat`. The batch file also installs any missing dependencies.

## Try the sample workflow

1. Click **Open Excel** and select `examples/sample_records.xlsx`. The sample contains two fictitious records and no real phone numbers. Changes are saved to the selected workbook, so copy it first if you want to keep a pristine sample.
2. Click **Settings → Certificate Templates**. Select `examples/volunteer_sample.pptx`. Leave the name marker as `Name`. Above the volunteer table, select Participant Name, Date, and Certificate Number for the three certificate columns.
3. Select `powerpoint` (or `auto`) on Windows with PowerPoint installed. Without PowerPoint, select `libreoffice`; if necessary, enter `C:\Program Files\LibreOffice\program\soffice.exe` as the executable path.
4. Click **More Actions → Generate / Retry PDFs**. Volunteers should show **PDF ready**. Beneficiaries do not receive certificates.
5. Click **More Actions → Open PDF Folder** to inspect the output. Check fonts, layout, and long names before sending real certificates.
6. Type directly into one of the six empty rows at the bottom to enter another record. With **Save automatically** and **Create certificates automatically** checked, edits save about 1.4 seconds after typing stops and eligible certificates generate in the background.

Types are `volunteer` and `beneficiary`. Imported multilingual headings, names, and recognized type aliases are supported; user data is preserved as entered while application controls and messages are in English.

## Three workflow tabs

- **Volunteer Certificates**: volunteer records and PPTX certificate sharing.
- **Family Applications**: family responses and a complete confirmation PDF for the applicant to review.
- **Volunteer Activities**: volunteer responses and personalized activity messages.

Every tab uses the same Excel-style editing controls. Tabs filter the selected worksheet by record type; use the Worksheet selector to choose the relevant Google response or local worksheet. Each tab remembers its selected worksheet during the session. Volunteer certificate and activity tabs can deliberately use the same registrations, avoiding duplicated people.

Six empty entry rows remain at the bottom. Start typing in any of them to create a record; six new empty slots remain below it. The current tab supplies the type automatically when a Type column is mapped. Empty entry rows are not saved to Excel. If an imported sheet has no Type column, configure its default type in Worksheet Columns. Imported Google response rows remain read-only, but you can add manual records in the blank rows below them. Manual records are not sent back to Google.

**Family confirmation:** select the Family Applications tab, choose its synced response worksheet, then click the applicant's **Share** button. The app creates an A4 PDF containing every original Google question/answer (including unanswered fields), with a request to confirm accuracy or provide corrections. Review the WhatsApp message, click Open WhatsApp, and attach the PDF manually. Files are under `data/family_confirmations`; no PowerPoint or LibreOffice is needed for these forms. This is a review summary, not an approval or an interactive fillable PDF.

**Open Application Form** opens the supplied family form editor (requires the form owner's Google permissions):
https://docs.google.com/forms/d/1Z8bwgdGnaaTqHR5WCY-RQu0gD9v-o25xoUGByiqW2Ok/edit

A form editing URL is not a response spreadsheet URL. Open the form editor → Responses → green Google Sheets icon to obtain the linked Sheet URL. The owner has not provided that response Sheet or credentials yet, so live sync is not configured. Link the form to a Google Sheet and configure that response Sheet URL and service-account permissions in Google Sheets Connection. Set its type to beneficiary. The form content has not been verified; no fixed question list is assumed: all response-sheet headings and values are preserved dynamically.

**Volunteer activity messages:** configure a volunteer Google source and select it under Volunteer Activities. Click **Event Message** to save the event, date, time, location, details, and message template. Supported placeholders: `{name}`, `{phone}`, `{event}`, `{date}`, `{time}`, `{location}`, `{details}`. Each Share button fills the selected applicant's details and opens an editable preview. Cancel sends nothing; Open WhatsApp opens the selected phone number's chat for manual sending. Activity messages do not require a certificate PDF.

Settings and advanced actions remain under **Settings** and **More Actions**. All interface text is English; participant data keeps its original language.

## Excel records and column mapping

The default database is `data/community.xlsx`. **Open Excel** can select an existing `.xlsx` or `.xlsm` workbook; its first row is treated as column headings. **Settings → Worksheet Columns** selects the participant name, type, phone, attendance date, and certificate number columns. If a worksheet has no Type column, choose a default type for that worksheet.

Edit cells directly, press Tab to move between cells, search any field, switch workflow tabs, and paste multiple rows or columns from Excel with Ctrl+V. Ctrl+C copies selected cells. While a filter is active, paste fills existing visible rows; clear filters to paste additional rows.

This is a table editor, not a complete Excel replacement: it does not reproduce charts or merged-cell layouts or recalculate formulas. Unedited numeric, date, and formula cells retain their values. Newly entered text is stored as text rather than executed as a formula. Deleting rows may affect formula references or Excel tables.

The app adds hidden `__fb_id`, `__fb_source`, and `__fb_key` columns for stable record identity. Keep these columns and do not manually copy or change record IDs. Each actual write creates a backup in `<workbook_name>_backups`. Unchanged Google polling does not rewrite the workbook or create a backup. Backups do not expire automatically.

Saving uses a temporary file, atomic replacement, an application lock, and external modification checks. Use one administrator instance per workbook. If external changes conflict with unsaved edits, auto save pauses: copy your edits somewhere safe, then reload and resolve them. Close Excel if it locks the workbook. A stale `.foodbank.lock` file should only be removed after all app instances are closed.

## Certificate templates

The recipient name must be editable PPTX text, not text baked into an image. The default `Name` marker (and legacy `auto` setting) detects the standalone word `name` (case-insensitive) or `{{NAME}}`, including markers split across text runs. Alternatively, enter the exact text to replace. Text boxes, grouped shapes, table cells, and layout/master text are supported. Use the editable text **Date No** for attendance date and **Cert No** for certificate number. The main Volunteer Certificates screen has **Name column**, **Attendance date column**, and **Certificate number column** selectors. Each selection is saved per workbook/worksheet and applies to every record in that worksheet. Select columns, not individual row numbers.

Use the actual attendance date, not the form submission timestamp. The certificate number is read from the chosen column; the app does not invent or auto-increment numbers. Store numbers such as `0007` as Excel text to preserve leading zeros. If the template contains a date/number marker but the mapped value is blank, the app reports the missing field and does not export an incomplete certificate. Changing a date, number, or column selection invalidates the cached PDF.

For local Excel sheets, a missing Date or Certificate Number column is offered and saved on the next record save. For Google imports, select the appropriate `Google · <heading>` column; maintain these values in the source Google Sheet and sync again, since imported rows are read-only. All original source columns are available for selection.

Legacy `{{DATE}}`, `{{TYPE}}`, and `{{ID}}` markers still work; `{{CERT_NO}}` is also supported. `{{ID}}` is the internal record ID, not the administrator's certificate number.

The original template is never edited. A temporary PPTX is filled and converted to a PDF under `data/certificates/volunteer`. Generation is checked at startup, after saved edits, when external Excel changes are loaded, and after Google sync, across all worksheets. Matching cached PDFs are reused; changes to names, dates, certificate numbers, types, templates, or converter selection cause regeneration on the next check. Click **More Actions → Reload Excel** or **More Actions → Generate / Retry PDFs** after modifying a template file.

- `templates/volunteer_template.pptx` is the supplied original volunteer template.
- `examples/volunteer_sample.pptx` is a clearly labeled demonstration with all three markers.
- Windows `auto` uses PowerPoint COM and requires PowerPoint and pywin32. On Linux/macOS, `auto` uses LibreOffice Impress.
- Missing templates/converters and export errors appear in the certificate status. Fix the prerequisite and click **More Actions → Generate / Retry PDFs**.
- Long names may overflow the original design; enlarge the text box or adjust template font sizes. LibreOffice and PowerPoint may render fonts and line breaks differently.
- Historical PDFs are retained. The Share action uses only a PDF whose signature matches the current record and template.

## Connect Google Forms / Google Sheets

1. Link Google Form responses to a Google Sheet. For Tally volunteer registrations, use Tally → Integrations → Google Sheets, then connect that response Sheet here. No direct Tally API is used.
2. Enable the **Google Sheets API** in Google Cloud, create a service account, and download its JSON key to a secure local location.
3. Share each Google Sheet with the JSON file's `client_email`, granting **Viewer** access. Never commit this key to GitHub.
4. Open **Settings → Google Sheets Connection**, select the credentials JSON, and add a source for each registration category. Enter the type, Sheet URL/ID, A1 range including its header row (for example `'Form Responses 1'!A1:ZZ`), and a separate local destination worksheet name (maximum 31 characters).
5. Enter the exact source headings for the name, phone, email, and date columns. Name is required; the others are optional. A unique, nonempty response ID column is strongly recommended.
6. Enable **Auto sync** (default interval: 60 seconds), or click **Sync Google** manually.

Source fields are retained as `Google · <original heading>` columns alongside normalized columns. Each source owns a separate local worksheet, whose imported response rows are read-only in the app. Manual records can be entered in the empty rows below.

Sync is **Google → local Excel only**; it never writes to Google. Stable response IDs update existing records without duplicating them. Records removed from Google are not automatically deleted from local history. New volunteer responses trigger certificate checks; beneficiaries are synced without certificates and can receive application confirmation PDFs from the Family Applications tab.

**Without a unique ID column, identity is based on Google row numbers: sources must be append-only. Do not sort, insert, or delete existing response rows.** Use a new destination worksheet when changing a source, registration type, or identity-column scheme. A timestamp is suitable only if it is truly unique.

Real access requires your credentials and sharing permissions. Network access to `oauth2.googleapis.com` and `sheets.googleapis.com` is required. Failures appear in the status area and automatic sync retries on the next interval; local Excel editing remains available.

## WhatsApp sharing

The last table column contains **Share**. Clicking it opens the recipient's chat. If a current PDF is ready, the app also opens its folder and copies the PDF path. **Attach the PDF and send manually in WhatsApp.** Ordinary WhatsApp links cannot attach local files automatically; no messages are sent automatically.

Default country code: Malaysia `60`. Local numbers starting with `0` are normalized accordingly; `+60...` and `0060...` are supported. Change the default with **Settings → WhatsApp Country Code**. Store phone numbers as Excel text to preserve leading zeros.

## Updating from the previous version

The donor tab, template configuration, and new donor sources have been removed. Existing donor Excel rows and old PDFs are retained as historical data; donor sources saved in old settings are ignored and donor certificates are no longer generated. Back up `data`, your Excel workbook, and custom PPTX templates before replacing program files. Existing local worksheets gain the certificate number column when saved.

## Settings and recovery

Settings and the certificate index are stored in `data/settings.json` and `data/certificate_index.json`. Set `FOODBANK_DATA_DIR` to use another data directory. Back up that directory and the selected workbook. Corrupt settings/index JSON raises an error instead of silently overwriting it. Closing prompts for unsaved changes and waits for active background tasks; LibreOffice export has a 120-second timeout.

## Linux / macOS and development

Install Python and LibreOffice Impress, then run `bash run.sh`, or create a virtual environment and install `requirements.txt` manually. A desktop display is needed for interactive use. Headless validation uses Qt's offscreen backend:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m unittest discover -s tests -v
# Include real PDF generation and name verification (requires poppler-utils):
QT_QPA_PLATFORM=offscreen FOODBANK_TEST_SOFFICE=/usr/bin/libreoffice .venv/bin/python -m unittest discover -s tests -v
```

Google tests use mocked API responses; they do not verify your live account. WhatsApp tests intercept browser calls and send no messages. Windows PowerPoint COM export must be checked on the target machine.

Code: `applications.py` (family confirmation PDFs and activity messages), `main.py` (English desktop UI/background tasks), `storage.py` (Excel), `certificates.py` (PPTX/PDF/share URLs), `google_sync.py` (Google read adapter), and `tests/`. `tools/create_examples.py` regenerates fictitious examples; `tools/package_source.py` packages the allowlisted source/assets without live records, credentials, caches, or virtual environments. This is Python source, not a Windows executable installer.
