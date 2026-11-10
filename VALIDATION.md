# Validation

Environment: Linux, Python 3.12, PySide6 6.12, openpyxl 3.1.5, python-pptx 1.0.2, LibreOffice Impress 25.2.

Command:

```bash
QT_QPA_PLATFORM=offscreen FOODBANK_TEST_SOFFICE=/workspace/foodbank-libreoffice/soffice \
  /workspace/foodbank-venv/bin/python -m unittest discover -s tests -v
```

Result: **42 tests passed, 0 skipped**.

Coverage:

- Three workflow tabs, continuous empty rows, type defaults, blank-row omission, family PDF field preservation, activity-template validation and cancelled-share behavior; legacy entry form saves a real Excel record; settings and advanced actions remain reachable through their menus.

- English controls, dialog titles, labels, headers, validation errors, and background messages. Chinese input aliases remain supported without becoming interface labels.
- English type filtering preserves multilingual participant data.
- Desktop startup, settings dialogs, multi-row paste, auto save, search, and reload.
- Excel stable IDs, duplicate ID repair, preservation of formulas/numbers/other worksheets, backups, deletion, and external-change conflict protection.
- Mocked Google API column mapping, original field retention, stable IDs, idempotent sync, updates, history retention, worksheet conflicts, and no rewrite on unchanged polling.
- Main-screen name/date/number mapping persists per worksheet; adding the certificate number column to legacy workbooks preserves record IDs.
- Date No and Cert No split-run replacement, leading-zero certificate numbers, literal marker-like recipient values, missing-field blocking, and number-driven cache invalidation.
- PPTX split-run markers, custom markers, table-cell markers, missing marker/template handling, cached exports, and regeneration after record changes.
- Real LibreOffice PDF conversion with pdftotext verification of recipient names, attendance dates, and certificate numbers.
- Background certificate generation across worksheets, using the supplied volunteer template. Legacy donor rows remain stored but never generate certificates. Beneficiaries receive no certificate.
- WhatsApp browser and folder calls intercepted in tests; no messages sent.

`examples/interface.png` shows the English application with fictitious records after a real volunteer PDF export. File pickers use Qt dialogs so their controls do not inherit a non-English native OS dialog.

Not verified: a live Google account/sheet (credentials and sharing permissions were not supplied), Windows PowerPoint COM conversion, live WhatsApp login/sending. Check fonts, long names, and approved template layout on the target Windows computer before production use.

Live Forms/Sheets sync is not verified without the response Sheet URLs and authorized credentials. The supplied family Form edit link is wired to Open Application Form; the owner has not supplied its response Sheet yet. This cloud network does not permit docs.google.com, so the form content has not been inspected. Confirmation PDFs use the complete response-sheet fields rather than an assumed form schema. Google Form URLs entered as Sheet URLs receive an actionable validation error. Tally registrations connect through their linked Google Sheet; direct Tally API access is not implemented.
