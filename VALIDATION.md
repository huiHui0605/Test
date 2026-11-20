# Validation

Validated 2026-11-20 on Debian 13, Python 3.12, PySide6 6.12, openpyxl 3.1.5, python-pptx 1.0.2, google-auth-oauthlib 1.5.0, keyring 25.7.0 and LibreOffice Impress 25.2.3.2.

Starting revision for PDF folder detection: `7f7ec9b050ef2e6e026428bffc83fd2a5da13b3d` on `huiHui0605/Test/main`. The remote was fetched and checked before implementation. The original `Foodbank-System` repository was not modified.

## PDF folder detection and reuse

- Generation and manual sharing check the actual file, record/template signature and available PDF checksum. Matching PDFs are skipped without conversion or PDF modification; missing, changed or outdated files are not marked ready.
- New `.pdf.json` receipts support recovery without the central UI index and after moving the complete output folder. Existing matching legacy index entries remain supported; unknown files are never associated by name alone.
- Certificate cells now show `PDF ready` / `PDF not ready`, with details in their tooltip. The existing 2.5-second monitor and new Check PDF Folder action refresh the current worksheet even with auto-generation disabled, updating only status cells during editing.
- Added seven tests covering cacheless recovery and unchanged PDF mtime, deletion/restoration, altered/empty files, changes to record/template, same-name different-ID records, legacy adoption, unknown PDFs, folder relocation, repeat manual generation, missing-file sharing, recovery after table reload, and preserving unsaved edits/error explanations.
- The full suite still includes real LibreOffice PDF generation for multiple rows, including datetime formatting and independent names/dates/numbers. New corruption and conversion-count scenarios use local fake converter output; no real messaging occurs.

## Attendance date/time formatting follow-up

- Attendance values loaded as Excel datetime objects or ISO datetime strings now display/export only their date. The generic storage serialization and unrelated timestamps are unchanged.
- Certificate numbers stored as text remain exact, including leading zeros and slashes. Date-typed numbers block generation with instructions to re-enter the intended identifier as text; no attempt is made to guess the original identifier.
- The real three-person desktop PDF test now starts with native Excel datetime cells, checks each date/number, and rejects the `T00:00:00` suffix. The standalone real PDF test also uses a datetime object.
- New tests cover date/date-time formatting, preservation of literal identifiers, date-typed number rejection, table display, unchanged native source cells after editing another field, and correction of a number through the local table editor.

## Per-person certificate regression

The bundled production PPTX had editable `Name` but literal `July 18, 2026` and `26/06/3705`. The row mapping was already per-person; those two literal template values were never replaced. Earlier PDF tests checked the bundled template's name only, while date/number tests used a synthetic template, so they did not catch this asset defect.

- Changed only those two text values to `Date No` and `Cert No`. ZIP-entry comparison confirms only `ppt/slides/slide1.xml` changed; all other original template assets are byte-identical.
- Template settings and generation now reject a missing name/date/number marker. Missing-marker checks run before cache reuse; cached PDF signatures were versioned to regenerate older output.
- The real desktop background test now exports three rows using the bundled template, custom column mappings, deliberately incorrect unselected columns, distinct dates/numbers, and two people with the same name. Extracted PDF text matches each row's three values, excludes the other rows' dates/numbers and old literal values, and three distinct files exist.
- Added missing-marker/cache rejection, template-dialog validation, `Attend Date` column recognition, and two-person bundled-template replacement tests. Leading-zero numbers remain intact.
- Separately generated fictitious Alice/Bob PDFs, extracted their dates/numbers, and visually inspected Alice's rendered certificate. The existing design and field positions remain intact. Windows PowerPoint rendering still requires target-computer verification.

## Automated result

**92 tests passed, 0 failed, 0 skipped**, including real LibreOffice conversion.

```bash
QT_QPA_PLATFORM=offscreen XDG_CACHE_HOME=/workspace/work/cache \
  FOODBANK_TEST_SOFFICE=/workspace/work/soffice \
  /workspace/work/foodbank-venv/bin/python -m unittest discover -s tests -v
```

`/workspace/work/soffice` is a task-local launcher for verified Debian LibreOffice packages extracted into the workspace, with relocated paths and libraries. It is a validation-environment artifact, not required by the application and not distributed. On a normal Linux installation use `FOODBANK_TEST_SOFFICE=/usr/bin/libreoffice`. On Windows install PowerPoint or LibreOffice normally.

Also checked: Python compilation, `git diff --check`, allowlisted ZIP integrity and module coverage, source manifest, and tracked-file/ZIP checks for runtime data and credentials. The curated source ZIP contains all new modules and the Chinese guide. The GitHub revision ZIP contains the committed source/assets, not a Python environment or Windows executable.

## Existing behavior retained

- Three English workflow tabs; six empty entry rows; type defaults and blank-row omission; paste, autosave, search and reload.
- Excel stable IDs, migration of certificate-number columns without losing IDs, duplicate-ID repair, formulas/numbers, other worksheets, backups, deletion and external-change conflict protection.
- Name / Date No / Cert No replacement across PPTX runs; leading-zero certificate numbers; literal marker-like values; required-field blocking; cache invalidation on changed records/columns/templates.
- Real LibreOffice PDF text checks for recipient names, dates and certificate numbers; background generation across worksheets and supplied volunteer template.
- No donor workflow/certificate generation; historical donor rows retained.
- Google merge identity, original answers, idempotence, updates, history retention and destination ownership and case-collision protection. Unchanged polling does not rewrite Excel.

## New behavior validated with mocks/fictitious data

- OS credential-store read/write/delete, fail-closed storage behavior and sanitized failures; credential refresh and revoked authorization; missing permissions; Desktop App configuration/official endpoint checks.
- Loopback/PKCE login configuration, partial-consent rejection, cancelled login and account-identity lookup failure leaving previous credentials intact; token material is held only by the mocked secure store in tests.
- Paginated Drive spreadsheet metadata listing; selected worksheet headers and escaped titles; both source panels, field dropdowns, saved choices and cancellation; explicit required fields/row-number fallback and rejection of Timestamp for volunteer attendance.
- Stable numeric worksheet ID after rename, stable registration ID after reordering, one-row updates without duplication, duplicate-ID rejection and original-question retention.
- Different activity dates for different volunteers; legacy global date ignored; certificate date mapping cannot override the Google participation date; missing/invalid phone/date and missing `{date}` template marker block sharing.
- Full original application answers in a real Qt-generated PDF, blank-answer labeling and the email confirmation/non-approval wording.
- Exact single applicant recipient, header-injection rejection, real PDF bytes and MIME attachment filename; Gmail `userId=me`, no HTTP/auth or application send retries. A real Google Gmail client is exercised against a mocked requests transport; a simulated dropped connection verifies a single request with retries and redirects disabled. This avoids httplib2’s internal stale-connection replay.
- Preview cancellation sends nothing and creates no delivery reservation; missing email blocks the action; review checkbox gates Send Email; previous attempts require an additional explicit resend acknowledgement.
- Duplicate clicks blocked; durable journal reservations reject stale reviews; successful/failed/ambiguous responses, timeout/server failure, missing response ID, credential failure and journal-write failure handled conservatively.
- Interrupted attempts remain uncertain after reopening the journal. The app does not claim Gmail exactly-once delivery and requires a human Sent-mail check before resending.
- A deliberately waiting background Google task leaves the Qt event loop responsive. Disconnect removes the mocked authorization and disables the sync timer while retaining Excel history.

Some former family-WhatsApp/global-date tests were updated for the explicitly requested email/per-person-date behavior; certificate tests now also enforce all three required per-person fields on the bundled template. Google and Gmail calls are mocked; browser/desktop-share calls are intercepted in automated tests. No real email or WhatsApp message was sent.

## Visual and document review

`tools/capture_examples.py` uses fictitious data, mocked Google listings and an intercepted/cancelled email preview. It generates actual Qt screenshots, not illustrative mockups:

- `examples/interface.png`: the three tabs, mapping dropdowns, six blank rows and a real volunteer certificate marked PDF ready, with `0007`.
- `examples/google_selection.png`: the two-source picker and participation-date mapping.
- `examples/family_email_review.png`: applicant email, message, PDF preview and disabled Send Email until reviewed.
- `examples/family_confirmation_demo.pdf`: only Example Family, `family@example.test`, `DEMO-IC-NOT-REAL` and other fictitious answers.

The screenshots were opened and visually inspected. The family PDF was rendered and visually inspected; `pdftotext` confirmed the complete demo fields, multi-line answer, blank answer and non-approval message. No real applicant data, account credential or token is present in these assets.

The [Tally official integration guide](https://tally.so/help/google-sheets-integration) was fetched and read on 2026-11-16. The Chinese guide documents the current native integration path and **Export existing submissions** switch without assuming it was enabled for the organization.

## Still requires target-computer acceptance

- Actual OAuth client/consent configuration, Google account login, real secure credential-store persistence and browser loopback callback on the target Windows computer.
- Organization test users/production verification, scopes, Workspace policy, real Sheet access and the two selected response worksheets. No real source Sheet was supplied or connected.
- Live Gmail submission and delivery to the responsible person's own test mailbox, including PDF attachment checks, quota/policy conditions and manual Sent-mail verification.
- Windows PowerPoint COM conversion, template fonts/long names, and live WhatsApp login/manual sending.

The supplied family Form URL is an editor link. Its private questions and corresponding response Sheet have not been inspected or inferred. Tally data is accessed only through its connected Google Sheet. Complete the acceptance checklist in [OPERATOR_GUIDE_ZH.md](OPERATOR_GUIDE_ZH.md) before production use.
