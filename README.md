# Ted Support Assistant: local pilot

This prototype addresses the team's generic-document intake and FAQ pilot. It searches an Excel workbook by keyword, creates local mock PBIs for demo-validated intake requests, saves documentation-gap reports to a local SQLite database, and can record a **mock** Halo incident.

The visual theme is Rabobank-inspired; this is an unofficial prototype, not an official Rabobank service.

**The workbook and form contain demonstration guidance, not approved TED policy.** PDF/A, malware, DLP and service-account permission checks are simulated, not performed. Duplicate checks only cover local mock submissions. There is no real Halo/Azure DevOps or SharePoint connection, or notification to TED. Do not enter real customer information, tokens or secrets.

## Setup

Requires Python 3.9 or newer with `venv` and `pip` available:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
```

On Windows, use `.venv\Scripts\python` in place of `.venv/bin/python`. Open <http://127.0.0.1:8000>. The server only listens on this computer; it is not a shared team deployment.

Optionally set `PORT`, `TED_DB_PATH` (SQLite database path) or `TED_FAQ_PATH` (workbook path) before starting. Defaults are port `8000`, `data/reports.sqlite3`, and `data/ted_faq.xlsx`. The database is created on first start; existing documentation reports are preserved.

## Maintain the FAQ

Open `data/ted_faq.xlsx` in Excel and edit the `FAQ` sheet. Keep the header row exactly as it is:

| Column | Contents |
| --- | --- |
| `id` | Unique identifier using letters, numbers, `_` or `-` |
| `topic` | Subject or category |
| `question` | Common user question or action |
| `answer` | Reviewed response |
| `keywords` | Search terms and synonyms |
| `reference` | Source name or an HTTP(S) URL; URLs appear as links |
| `owner` | Person or team responsible for the answer |
| `last_reviewed` | Review date or review status |

All columns are required; leave unused rows completely blank. Do not use formulas. Save and close the workbook after editing. The next search reads the saved file; no server restart is needed. If the workbook is missing or invalid, the search page returns an error and details are written to the server log rather than silently using stale data. Keep a backup before editing.

The six included entries are **unapproved examples**. TED should replace them with reviewed answers and authoritative source references before any real use. Try `PDF/A`, `document keys`, or `API onboarding`. Search is keyword matching, not an AI model.

## Test the intake and reports

1. Select **Start a generic-document request** and download the sample template and PDF. Complete the form with fictional details, attach either file or both, and choose a simulated service outcome. If attaching Excel, its single data row must match the form. For the prefilled sample, use Add, `sample.pdf`, English, Demo document and product types, the template's date, and key `DEMO-001`. Set a requester and environment.
2. The demo checks the required fields, replacement old key, valid non-past effective date and duplicate new keys in the local mock database. For an attached Excel file, it checks the demo template version and columns and matching spreadsheet and form values. For an attached PDF, it checks the filename, basic readability and duplicate PDFs. Each file is limited to 5 MB. Checks requiring an absent file, including the PDF/A simulation when no PDF is provided, are shown as skipped. PDF/A compliance is not verified; service outcomes are selected simulations, not security scans. Files are processed in memory and not retained.
3. If applicable demo checks and simulated services pass, a local `DEMO-<id>` mock PBI is saved, even when checks were skipped for a missing file. This is not a fully validated submission. Failures and intentional duplicate or legacy exceptions are rejected; no manual-review workflow exists yet. The same key or PDF cannot be reused in another mock submission. No real PBI is created.
4. Search for a missing topic and submit the documentation-gap form. Optionally select **Create a mock Halo incident**. A local `MOCK-<report id>` is shown; nothing is sent to Halo.
5. Use a SQLite client to inspect `mock_pbis` and `reports` in `data/reports.sqlite3`. Earlier `document_requests` records are preserved but no longer receive new entries. These are local-only records and do **not** create work for TED.

Run automated tests:

```sh
.venv/bin/python -m unittest discover -s tests -v
```

## Next steps with TED

1. Assign an FAQ owner and have TED review the workbook, add authoritative references, and agree the real intake fields, checklist and exceptions.
2. Test the prototype with five requesters using fictional data. Baseline clarification rounds and preventable rejections before setting improvement targets.
3. Agree whether Halo or Azure DevOps is the system of record, the approval process, and the owner of the queue. Integrate only after access and security review.
4. Design file storage and malware/DLP scanning, then implement and test real template, key, filename, date and PDF/A checks; provide a manual exception route.
5. Prioritize onboarding/status, subscription approval, alert routing/deduplication and consumer outreach separately. Preserve the existing urgent on-call path.

For real shared use, add authentication and authorization, CSRF protection, secure hosting, access-controlled storage, retention policies, and privacy/security review. The current localhost-only prototype has none of these protections.
