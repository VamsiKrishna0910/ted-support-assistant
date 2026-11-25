# Ted Support Assistant: local pilot

This prototype addresses the team's generic-document intake and FAQ pilot. It searches an Excel workbook by keyword, saves guided intake requests and documentation-gap reports to a local SQLite database, and can record a **mock** Halo incident.

**The workbook and form contain demonstration guidance, not approved TED policy.** There is no file upload, PDF/A verification, duplicate-key lookup, automated validation of the actual documents, real Halo/Azure DevOps connection, or notification to TED. Do not enter real customer information, tokens or secrets.

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

1. Select **Start a generic-document request**, enter fictional details, review the manual checklist and save. Replacement requests require an old key. The date format is checked, but the application cannot verify the document, keys, filename or date against real systems.
2. Search for a missing topic and submit the documentation-gap form. Optionally select **Create a mock Halo incident**. A local `MOCK-<report id>` is shown; nothing is sent to Halo.
3. Use a SQLite client to inspect the `document_requests` and `reports` tables in `data/reports.sqlite3`. These are local-only records; saving them does **not** create work for TED. Existing `reports` are not modified by the intake feature.

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
