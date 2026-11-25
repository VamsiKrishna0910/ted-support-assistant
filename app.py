"""Local TED support pilot with an Excel FAQ and guided intake."""

from datetime import date, datetime
from contextlib import closing
from email import policy
from email.parser import BytesParser
import hashlib
import html
from io import BytesIO
import os
from pathlib import Path
import re
import sqlite3
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlencode, urlsplit
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile, ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError


ROOT = Path(__file__).resolve().parent
FAQ_PATH = ROOT / "data" / "ted_faq.xlsx"
DEFAULT_DB = ROOT / "data" / "reports.sqlite3"
MAX_BODY = 16_384
MAX_FILE_SIZE = 5 * 1024 * 1024
MAX_UPLOAD_BODY = 2 * MAX_FILE_SIZE + 16_384
MAX_WORKBOOK_EXPANDED = 25 * 1024 * 1024
DEMO_TEMPLATE_VERSION = "DEMO-1"
DEMO_COLUMNS = ("template_version", "filename", "language", "document_type", "product_type", "effective_date", "new_key", "old_key")
FAQ_COLUMNS = ("id", "topic", "question", "answer", "keywords", "reference", "owner", "last_reviewed")
STOP_WORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "can", "do", "does",
    "for", "how", "i", "in", "is", "it", "me", "of", "on", "or", "the",
    "to", "what", "when", "where", "which", "with",
}


def load_faq(path):
    """Read the workbook on each request so saved changes become searchable."""
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if "FAQ" not in workbook.sheetnames:
            raise ValueError("Workbook must contain an FAQ sheet.")
        rows = workbook["FAQ"].iter_rows()
        header = next(rows, None)
        if header is None or tuple(cell.value for cell in header) != FAQ_COLUMNS:
            raise ValueError("FAQ columns do not match the required headers.")
        entries = []
        ids = set()
        for row_number, cells in enumerate(rows, start=2):
            values = [cell.value for cell in cells]
            if not any(value is not None for value in values):
                continue
            if len(values) != len(FAQ_COLUMNS) or any(cell.data_type == "f" for cell in cells):
                raise ValueError(f"FAQ row {row_number} has extra columns or formulas.")
            if any(not isinstance(value, (str, date, datetime)) or not str(value).strip() for value in values):
                raise ValueError(f"FAQ row {row_number} must have a value in every column.")
            entry = {key: str(value).strip() for key, value in zip(FAQ_COLUMNS, values)}
            if not re.fullmatch(r"[a-zA-Z0-9_-]+", entry["id"]) or entry["id"] in ids:
                raise ValueError(f"FAQ row {row_number} has an invalid or duplicate id.")
            ids.add(entry["id"])
            entries.append(entry)
        if not entries:
            raise ValueError("FAQ sheet must contain at least one entry.")
        return entries
    finally:
        workbook.close()


def words(text):
    return {word for word in re.findall(r"[a-z0-9]+", text.lower()) if word not in STOP_WORDS}


def search_faq(query, entries):
    terms = words(query)
    if not terms:
        return []
    results = []
    for entry in entries:
        searchable = words(" ".join(entry.values()))
        matches = len(terms & searchable)
        if matches:
            results.append((matches, entry))
    results.sort(key=lambda result: -result[0])
    return [entry for _, entry in results[:5]]


MOCK_PBI_SCHEMA = """CREATE TABLE IF NOT EXISTS mock_pbis (
    id INTEGER PRIMARY KEY,
    requester TEXT NOT NULL,
    operation TEXT NOT NULL,
    environment TEXT NOT NULL,
    filename TEXT NOT NULL,
    language TEXT NOT NULL,
    document_type TEXT NOT NULL,
    product_type TEXT NOT NULL,
    effective_date TEXT NOT NULL,
    new_key TEXT NOT NULL UNIQUE COLLATE NOCASE,
    old_key TEXT NOT NULL,
    pdf_sha256 TEXT UNIQUE,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
)"""


def initialize_db(path):
    with closing(sqlite3.connect(path)) as connection, connection:
        connection.execute("BEGIN IMMEDIATE")
        connection.execute(
            """CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY,
                query TEXT NOT NULL,
                details TEXT NOT NULL,
                mock_incident INTEGER NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        connection.execute(
            """CREATE TABLE IF NOT EXISTS document_requests (
                id INTEGER PRIMARY KEY,
                requester TEXT NOT NULL,
                operation TEXT NOT NULL,
                environment TEXT NOT NULL,
                filename TEXT NOT NULL,
                language TEXT NOT NULL,
                document_type TEXT NOT NULL,
                product_type TEXT NOT NULL,
                effective_date TEXT NOT NULL,
                new_key TEXT NOT NULL,
                old_key TEXT NOT NULL,
                checklist_confirmed INTEGER NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        connection.execute(MOCK_PBI_SCHEMA)
        columns = list(connection.execute("PRAGMA table_info(mock_pbis)"))
        if any(column[1] == "pdf_sha256" and column[3] for column in columns):
            expected = ("id", "requester", "operation", "environment", "filename", "language", "document_type",
                        "product_type", "effective_date", "new_key", "old_key", "pdf_sha256", "created_at")
            if tuple(column[1] for column in columns) != expected:
                raise ValueError("Cannot migrate an unexpected mock_pbis schema.")
            connection.execute("ALTER TABLE mock_pbis RENAME TO mock_pbis_previous_schema")
            connection.execute(MOCK_PBI_SCHEMA)
            connection.execute("INSERT INTO mock_pbis SELECT * FROM mock_pbis_previous_schema")
            connection.execute("DROP TABLE mock_pbis_previous_schema")


def header():
    return """<header class="topbar"><a class="brand" href="/"><span class="brand-mark" aria-hidden="true">T</span><span class="brand-text"><strong>TED Support</strong><small>RABOBANK</small></span></a><span class="demo-badge"><span></span> LOCAL PROTOTYPE</span></header>"""


def alert(message, kind="error"):
    if not message:
        return ""
    role = "status" if kind == "success" else "alert"
    return f'<div class="alert {kind}" role="{role}">{html.escape(message)}</div>'


def reference_html(reference):
    parsed = urlsplit(reference)
    text = html.escape(reference, quote=True)
    if parsed.scheme in ("http", "https") and parsed.netloc:
        return f'<a href="{text}" target="_blank" rel="noopener noreferrer">{text}</a>'
    return text


def page(query="", results=None, notice="", error="", details="", mock_incident=False):
    escape = html.escape
    if query and results:
        cards = []
        for entry in results:
            cards.append(
                f'<article class="result"><div class="result-meta"><span>{escape(entry["id"])}</span>'
                f'<span class="dot">·</span><span>{escape(entry["topic"])}</span></div>'
                f'<h3>{escape(entry["question"])}</h3><p>{escape(entry["answer"])}</p>'
                f'<p class="faq-meta">Reference: {reference_html(entry["reference"])} · Owner: {escape(entry["owner"])}'
                f' · Last reviewed: {escape(entry["last_reviewed"])}</p></article>'
            )
        result_html = (
            '<div class="section-heading"><div><span class="eyebrow">YOUR RESULTS</span>'
            f'<h2>Answers from the FAQ</h2><p>Showing {len(results)} matching FAQ entries.</p>'
            '</div><span class="count">EXCEL FAQ</span></div>' + "".join(cards)
        )
    elif query:
        result_html = (
            '<div class="empty"><span class="eyebrow">NO MATCHES FOUND</span>'
            '<h2>We couldn’t find that in the FAQ.</h2>'
            '<p>Try different keywords, or report what’s missing below.</p></div>'
        )
    else:
        result_html = (
            '<div class="empty"><span class="eyebrow">READY WHEN YOU ARE</span>'
            '<h2>Answers start with a question.</h2>'
            '<p>Search the sample TED FAQ, or start a guided generic-document request.</p></div>'
        )
    checked = " checked" if mock_incident else ""
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ted Support Assistant</title><link rel="stylesheet" href="/style.css"></head>
<body>{header()}
<main><div class="hero"><div class="hero-copy"><div class="pill"><span>✦</span> YOUR TED SUPPORT COMPANION</div>
<h1>Answers when<br>you <em>need them.</em></h1>
<p>Find guidance for common questions and prepare your generic-document request, all in one place.</p>
<form class="search" action="/" method="get"><span class="search-icon" aria-hidden="true">⌕</span>
<label class="sr-only" for="query">Search the FAQ</label><input id="query" name="q" value="{escape(query, quote=True)}" placeholder="Ask about documents or onboarding..." maxlength="500" required>
<button type="submit">Search <span aria-hidden="true">→</span></button></form>
<div class="suggestions">POPULAR TOPICS <a href="/?q=PDF%2FA">PDF/A</a><a href="/?q=document+keys">Document keys</a><a href="/?q=API+onboarding">API onboarding</a></div>
<a class="intake-link" href="/intake">Start a generic-document request <span aria-hidden="true">→</span></a>
</div><div class="hero-guide"><span class="guide-label">HERE TO HELP</span><h2>A little clarity goes a long way.</h2><p>Three simple ways to get started:</p>
<div class="guide-step"><span>01</span><div><strong>Find an answer</strong><small>Search the sample FAQ by keyword.</small></div></div>
<div class="guide-step"><span>02</span><div><strong>Prepare a request</strong><small>Use the guided document checklist.</small></div></div>
<div class="guide-step"><span>03</span><div><strong>Spot a gap</strong><small>Record missing guidance locally.</small></div></div></div></div>
<div class="pilot-banner"><strong>Sample content only.</strong> FAQ answers and intake checks require TED review; do not use them as production policy. File checks are local and do not verify TED requirements.</div>
<div class="content"><section class="answers" aria-live="polite">{result_html}</section>
<aside class="report"><div class="report-icon">✳</div><span class="eyebrow">HELP US IMPROVE</span><h2>Something missing?</h2>
<p>Record a gap locally. Do not include secrets or customer data.</p>{alert(notice, "success")}{alert(error)}
<form method="post" action="/report"><label for="report-query">Your search</label><input id="report-query" name="query" value="{escape(query, quote=True)}" placeholder="e.g. document keys" maxlength="500" required>
<label for="details">What’s missing?</label><textarea id="details" name="details" placeholder="Tell us what you hoped to find..." maxlength="3000" required>{escape(details)}</textarea>
<label class="checkbox"><input type="checkbox" name="mock_incident" value="1"{checked}><span><strong>Create a mock Halo incident</strong><small>Demo only. Nothing is sent to Halo.</small></span></label>
<button class="report-button" type="submit">Report missing documentation <span aria-hidden="true">→</span></button></form>
<div class="report-foot">Saved locally on this computer; not sent to TED.</div></aside></div></main>
<footer><span><strong>TED Support</strong> &nbsp; / &nbsp; RABOBANK</span><span>UNOFFICIAL LOCAL PROTOTYPE · SAMPLE CONTENT</span></footer>
</body></html>"""


INTAKE_FIELDS = (
    ("requester", "Requester or team", 200),
    ("filename", "PDF filename", 200),
    ("language", "Document language", 100),
    ("document_type", "Document type", 100),
    ("product_type", "Product type", 100),
    ("effective_date", "Effective date", 10),
    ("new_key", "New document key", 200),
    ("old_key", "Old document key (required for a replacement)", 200),
)


def intake_page(values=None, error="", notice="", validation_results=None):
    values = values or {}
    escape = html.escape
    inputs = []
    for name, label, maximum in INTAKE_FIELDS:
        value = escape(values.get(name, ""), quote=True)
        required = "" if name == "old_key" else " required"
        input_type = "date" if name == "effective_date" else "text"
        inputs.append(f'<label for="{name}">{escape(label)}</label><input id="{name}" name="{name}" type="{input_type}" value="{value}" maxlength="{maximum}"{required}>')

    def select(name, label, options):
        rendered = '<option value="">Select an option</option>'
        for value, text in options:
            selected = " selected" if values.get(name) == value else ""
            rendered += f'<option value="{value}"{selected}>{text}</option>'
        return f'<label for="{name}">{label}</label><select id="{name}" name="{name}" required>{rendered}</select>'

    operation = select("operation", "Request type", (("add", "Add"), ("replace", "Replace")))
    environment = select("environment", "Environment", (("non-prod", "Non-PROD"), ("prod", "PROD")))
    checked = " checked" if values.get("checklist") == "1" else ""
    exception_checked = " checked" if values.get("exception") == "1" else ""
    scenario = select("scenario", "Simulated service outcome", (
        ("pass", "All simulated checks pass"), ("pdfa", "Simulate PDF/A failure"),
        ("malware", "Simulate malware detection"), ("dlp", "Simulate DLP failure"),
        ("permission", "Simulate service permission failure"),
    ))
    validation_html = ""
    if validation_results is not None:
        results = "".join(
            f'<li class="validation-item"><strong class="validation-{status}">{status.upper()}</strong> {escape(message)}</li>'
            for status, message in validation_results
        )
        validation_html = f'<div class="validation-results" role="status"><h3>Demo validation results</h3><ul>{results}</ul></div>'
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Generic-document intake | Ted Support Assistant</title>
<link rel="stylesheet" href="/style.css"></head><body>{header()}
<main class="source"><a class="back" href="/">← Back to search</a><span class="eyebrow">GUIDED PILOT INTAKE</span>
<h1>Generic-document request</h1>
<p class="intro">Use fictional data only. This is a local demonstration using assumed rules, not TED policy. No files are retained or sent to SharePoint, TED or Azure DevOps.</p>
{alert(notice, "success")}{alert(error)}
{validation_html}
<form class="intake-form" method="post" action="/intake" enctype="multipart/form-data">{operation}{environment}{"".join(inputs)}
<fieldset><legend>Sample attachments</legend><p>Attach an Excel workbook, a PDF, or both. Checks that need a missing file will be skipped. <a href="/sample-template.xlsx">Download the demonstration Excel template</a> or a <a href="/sample.pdf">sample PDF</a>. If using the template, complete its single row to match the details above. Each file must be 5 MB or smaller.</p>
<label for="workbook">Excel workbook (.xlsx)</label><input id="workbook" name="workbook" type="file" accept=".xlsx">
<label for="pdf">PDF document (.pdf)</label><input id="pdf" name="pdf" type="file" accept=".pdf"></fieldset>
<fieldset><legend>Simulated controls</legend><p>PDF/A, malware, DLP and service-account permission checks are simulated, not performed. Passing them does not mean files are safe or PDF/A compliant.</p>
{scenario}
<label class="checkbox"><input type="checkbox" name="exception" value="1"{exception_checked}><span>This is an intentional duplicate or legacy exception (rejected in this demo).</span></label>
<label class="checkbox"><input type="checkbox" name="checklist" value="1"{checked} required><span>I understand this is a demonstration with simulated controls.</span></label></fieldset>
<button class="report-button" type="submit">Validate and create mock PBI →</button></form></main></body></html>"""


def sample_template():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Document"
    sheet.append(DEMO_COLUMNS)
    sheet.append((DEMO_TEMPLATE_VERSION, "sample.pdf", "English", "Demo", "Demo", date.today().isoformat(), "DEMO-001", ""))
    stream = BytesIO()
    workbook.save(stream)
    workbook.close()
    return stream.getvalue()


def sample_pdf():
    writer = PdfWriter()
    writer.add_blank_page(width=200, height=200)
    stream = BytesIO()
    writer.write(stream)
    return stream.getvalue()


def validate_documents(files, values):
    results = []
    workbook_name, workbook_data = files.get("workbook", ("", b""))
    pdf_name, pdf_data = files.get("pdf", ("", b""))
    if "workbook" not in files:
        results.append(("skipped", "Excel template and spreadsheet checks were not performed; no Excel file was supplied."))
    elif not workbook_name.lower().endswith(".xlsx") or not workbook_data:
        results.append(("failed", "Excel: select a non-empty .xlsx file."))
    else:
        try:
            with ZipFile(BytesIO(workbook_data)) as archive:
                entries = archive.infolist()
                if len(entries) > 1000 or sum(entry.file_size for entry in entries) > MAX_WORKBOOK_EXPANDED:
                    raise ValueError("Workbook is too large when expanded.")
            workbook = load_workbook(BytesIO(workbook_data), read_only=True, data_only=False)
            try:
                if "Document" not in workbook.sheetnames:
                    raise ValueError("Missing Document sheet.")
                sheet = workbook["Document"]
                if sheet.max_row is None or sheet.max_column is None or sheet.max_row > 2 or sheet.max_column > len(DEMO_COLUMNS):
                    raise ValueError("Only one row and the expected columns are supported.")
                rows = list(sheet.iter_rows(min_row=1, max_row=3, max_col=len(DEMO_COLUMNS) + 1))
                if len(rows) < 2:
                    raise ValueError("Missing data row.")
                if tuple(cell.value for cell in rows[0]) != DEMO_COLUMNS + (None,):
                    raise ValueError("Invalid headers.")
                if (len(rows) > 2 and any(cell.value is not None for cell in rows[2])) or rows[1][-1].value is not None:
                    raise ValueError("Only one row is supported.")
                if any(cell.data_type == "f" for cell in rows[1]):
                    raise ValueError("Formulas are not supported.")
                cells = [cell.value for cell in rows[1][:-1]]
                if isinstance(cells[5], datetime):
                    cells[5] = cells[5].date().isoformat()
                elif isinstance(cells[5], date):
                    cells[5] = cells[5].isoformat()
                if any(not isinstance(value, str) or not value.strip() for value in cells[:-1]):
                    raise ValueError("Missing required values.")
                if cells[-1] is not None and not isinstance(cells[-1], str):
                    raise ValueError("Invalid old key.")
                row = dict(zip(DEMO_COLUMNS, (str(value).strip() if value is not None else "" for value in cells)))
            finally:
                workbook.close()
        except (OSError, ValueError, KeyError, TypeError, IndexError, BadZipFile, InvalidFileException, ParseError):
            results.append(("failed", "Excel: use the demo template with one complete data row and no formulas."))
        else:
            if row["template_version"] != DEMO_TEMPLATE_VERSION:
                results.append(("failed", f"Template version must be {DEMO_TEMPLATE_VERSION}."))
            else:
                results.append(("passed", "Demo template version and required spreadsheet fields match."))
            if any(row[key] != values[key] for key in DEMO_COLUMNS[1:]):
                results.append(("failed", "Spreadsheet values must match the request details."))
            else:
                results.append(("passed", "Spreadsheet values match the request details."))
    if "pdf" not in files:
        results.append(("skipped", "PDF readability and filename matching were not checked; no PDF was supplied."))
    elif not pdf_name.lower().endswith(".pdf") or not pdf_data:
        results.append(("failed", "PDF: select a non-empty .pdf file."))
    else:
        try:
            reader = PdfReader(BytesIO(pdf_data), strict=True)
            if reader.is_encrypted:
                raise ValueError("Encrypted PDF.")
            if len(reader.pages) < 1:
                raise ValueError("PDF has no pages.")
        except (OSError, ValueError, KeyError, IndexError, PdfReadError):
            results.append(("failed", "PDF: document could not be opened or is encrypted."))
        else:
            results.append(("passed", "PDF: document opened for a basic format check; this is not a PDF/A check."))
    if "pdf" in files:
        if pdf_name != values["filename"]:
            results.append(("failed", "Uploaded PDF filename must match the request."))
        else:
            results.append(("passed", "PDF filename matches the request."))
        if "workbook" not in files:
            results.append(("skipped", "PDF-to-Excel filename matching was not checked; no Excel file was supplied."))
    return results


def parse_document_upload(content_type, body):
    message = BytesParser(policy=policy.default).parsebytes(
        b"Content-Type: " + content_type.encode("ascii") + b"\r\nMIME-Version: 1.0\r\n\r\n" + body
    )
    if message.defects or not message.is_multipart():
        raise ValueError("Invalid multipart upload.")
    parts = list(message.iter_parts())
    if len(parts) > 20:
        raise ValueError("Too many form fields.")
    files = {}
    form = {}
    seen_files = set()
    field_names = {name for name, _, _ in INTAKE_FIELDS} | {"operation", "environment", "checklist", "exception", "scenario"}
    for part in parts:
        name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        data = part.get_payload(decode=True)
        if part.defects or part.get_content_disposition() != "form-data" or not isinstance(data, bytes):
            raise ValueError("Invalid document upload.")
        if name in ("workbook", "pdf") and filename is not None and name not in seen_files:
            seen_files.add(name)
            if len(data) > MAX_FILE_SIZE:
                raise ValueError("Each file must be 5 MB or smaller.")
            if filename or data:
                files[name] = (filename, data)
        elif name in field_names and filename is None and name not in form and len(data) <= 3000:
            form[name] = [data.decode("utf-8")]
        else:
            raise ValueError("Invalid or duplicate form field.")
    return form, files


def validate_intake(form):
    values = {name: form.get(name, [""])[0].strip() for name, _, _ in INTAKE_FIELDS}
    for name in ("operation", "environment", "checklist", "exception", "scenario"):
        values[name] = form.get(name, [""])[0].strip()
    errors = []
    for name, label, maximum in INTAKE_FIELDS:
        value = values[name]
        if (name != "old_key" and not value) or len(value) > maximum:
            errors.append(f"{label} is required and must be at most {maximum} characters.")
    if values["operation"] not in ("add", "replace"):
        errors.append("Choose add or replace.")
    if values["operation"] == "replace" and not values["old_key"]:
        errors.append("An old document key is required for a replacement.")
    if values["old_key"] and values["old_key"].casefold() == values["new_key"].casefold():
        errors.append("Old and new document keys must differ.")
    if values["environment"] not in ("prod", "non-prod"):
        errors.append("Choose an environment.")
    if values["checklist"] != "1":
        errors.append("Confirm that you understand the simulated controls.")
    if values["exception"] not in ("", "1"):
        errors.append("Invalid exception selection.")
    if values["exception"] == "1":
        errors.append("Duplicate and legacy exceptions are rejected in this demo.")
    if values["scenario"] not in ("pass", "pdfa", "malware", "dlp", "permission"):
        errors.append("Select a simulated service outcome.")
    if values["effective_date"]:
        try:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", values["effective_date"]):
                raise ValueError("Invalid date format")
            if date.fromisoformat(values["effective_date"]) < date.today():
                errors.append("Demo rule: the effective date cannot be before today.")
        except ValueError:
            errors.append("Enter a valid effective date (YYYY-MM-DD).")
    return values, errors


def make_handler(faq_path, database):
    class Handler(BaseHTTPRequestHandler):
        def send_content(self, status, content, content_type="text/html; charset=utf-8", headers=None):
            body = content if isinstance(content, bytes) else content.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("X-Content-Type-Options", "nosniff")
            for key, value in (headers or {}).items():
                self.send_header(key, value)
            self.end_headers()
            self.wfile.write(body)

        def faq(self):
            try:
                return load_faq(faq_path)
            except (OSError, ValueError, KeyError, BadZipFile, InvalidFileException, ParseError) as exc:
                self.log_error("Could not load FAQ workbook: %s", exc)
                self.send_content(503, "FAQ workbook is unavailable or invalid. Check the server logs.", "text/plain; charset=utf-8")
                return None

        def do_GET(self):
            parsed = urlsplit(self.path)
            params = parse_qs(parsed.query)
            if parsed.path == "/":
                entries = self.faq()
                if entries is None:
                    return
                query = params.get("q", [""])[0].strip()[:500]
                submitted = params.get("submitted", [""])[0]
                notices = {
                    "report": "Your documentation report was saved locally; it was not sent to TED.",
                    "incident": "Your report and mock Halo incident were saved locally. Nothing was sent to Halo.",
                }
                incident_id = params.get("incident", [""])[0]
                if submitted == "incident" and re.fullmatch(r"[1-9][0-9]*", incident_id):
                    notices["incident"] = f"Your report and mock Halo incident MOCK-{incident_id} were saved locally. Nothing was sent to Halo."
                self.send_content(200, page(query, search_faq(query, entries), notices.get(submitted, "")))
            elif parsed.path == "/intake":
                self.send_content(200, intake_page())
            elif parsed.path == "/sample-template.xlsx":
                self.send_content(200, sample_template(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": 'attachment; filename="ted-demo-template.xlsx"'})
            elif parsed.path == "/sample.pdf":
                self.send_content(200, sample_pdf(), "application/pdf",
                                  {"Content-Disposition": 'attachment; filename="sample.pdf"'})
            elif parsed.path == "/style.css":
                self.send_content(200, (ROOT / "style.css").read_text(encoding="utf-8"), "text/css; charset=utf-8")
            else:
                self.send_content(404, "Not found", "text/plain; charset=utf-8")

        def do_POST(self):
            route = urlsplit(self.path).path
            if route not in ("/report", "/intake"):
                self.send_content(404, "Not found", "text/plain; charset=utf-8")
                return
            try:
                length = int(self.headers.get("Content-Length", ""))
            except ValueError:
                self.send_content(400, "A valid Content-Length is required", "text/plain; charset=utf-8")
                return
            limit = MAX_UPLOAD_BODY if route == "/intake" else MAX_BODY
            if length < 0 or length > limit:
                self.send_content(413, "Request body too large", "text/plain; charset=utf-8")
                return
            if route == "/intake":
                if self.headers.get_content_type() != "multipart/form-data":
                    self.send_content(415, "Unsupported content type", "text/plain; charset=utf-8")
                    return
                try:
                    form, files = parse_document_upload(self.headers.get("Content-Type", ""), self.rfile.read(length))
                except (ValueError, UnicodeError) as exc:
                    self.send_content(400, intake_page(error=str(exc)))
                    return
                self.save_intake(form, files)
                return
            if self.headers.get_content_type() != "application/x-www-form-urlencoded":
                self.send_content(415, "Unsupported content type", "text/plain; charset=utf-8")
                return
            try:
                form = parse_qs(self.rfile.read(length).decode("utf-8"), keep_blank_values=True, max_num_fields=20)
            except (UnicodeDecodeError, ValueError):
                self.send_content(400, "Invalid form data", "text/plain; charset=utf-8")
                return
            self.save_report(form)

        def save_intake(self, form, files):
            values, errors = validate_intake(form)
            if not files:
                errors.append("Select at least one Excel workbook or PDF.")
            if errors:
                self.send_content(400, intake_page(values, " ".join(errors) + " Please reselect your attachments."))
                return
            results = validate_documents(files, values)
            for key, label in (("pdfa", "PDF/A"), ("malware", "Malware"), ("dlp", "DLP"), ("permission", "Service-account permissions")):
                if key == "pdfa" and "pdf" not in files:
                    results.append(("skipped", "PDF/A simulation was not run; no PDF was supplied."))
                    continue
                status = "failed" if values["scenario"] == key else "simulated"
                results.append((status, f"{label}: simulated {'failure' if status == 'failed' else 'pass'}; no real service was called."))
            if any(status == "failed" for status, _ in results):
                self.send_content(400, intake_page(values, "Demo validation failed. No PBI was created; reselect your attachments to retry.", validation_results=results))
                return
            pdf_hash = hashlib.sha256(files["pdf"][1]).hexdigest() if "pdf" in files else None
            try:
                with closing(sqlite3.connect(database)) as connection, connection:
                    cursor = connection.execute(
                        """INSERT INTO mock_pbis
                        (requester, operation, environment, filename, language, document_type, product_type,
                         effective_date, new_key, old_key, pdf_sha256)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (values["requester"], values["operation"], values["environment"], values["filename"],
                         values["language"], values["document_type"], values["product_type"],
                         values["effective_date"], values["new_key"], values["old_key"], pdf_hash),
                    )
                    pbi_id = cursor.lastrowid
            except sqlite3.IntegrityError:
                results.append(("failed", "A mock submission already uses this key or PDF file. Duplicates are rejected."))
                self.send_content(400, intake_page(values, "No PBI was created; reselect your attachments to retry.", validation_results=results))
                return
            except sqlite3.Error as exc:
                self.log_error("Could not save mock PBI: %s", exc)
                self.send_content(500, intake_page(values, "Could not save the mock PBI. Please try again and reselect your attachments.", validation_results=results))
                return
            results.append(("passed", "No duplicate key was found in the local mock submissions."))
            if pdf_hash:
                results.append(("passed", "No duplicate PDF was found in the local mock submissions."))
            else:
                results.append(("skipped", "Duplicate PDF check was not performed; no PDF was supplied."))
            incomplete = " Some checks were skipped because an attachment was missing." if any(status == "skipped" for status, _ in results) else ""
            notice = f"Mock PBI DEMO-{pbi_id} created locally.{incomplete} Nothing was sent to Azure DevOps or TED; no files were retained."
            self.send_content(200, intake_page(notice=notice, validation_results=results))

        def save_report(self, form):
            query = form.get("query", [""])[0].strip()
            details = form.get("details", [""])[0].strip()
            mock_incident = form.get("mock_incident", [""])[0] == "1"
            if not query or not details or len(query) > 500 or len(details) > 3000:
                entries = self.faq()
                if entries is None:
                    return
                error = "Enter a search (up to 500 characters) and details (up to 3,000 characters)."
                self.send_content(400, page(query, search_faq(query, entries), error=error, details=details, mock_incident=mock_incident))
                return
            try:
                with closing(sqlite3.connect(database)) as connection, connection:
                    cursor = connection.execute(
                        "INSERT INTO reports (query, details, mock_incident) VALUES (?, ?, ?)",
                        (query, details, int(mock_incident)),
                    )
                    report_id = cursor.lastrowid
            except sqlite3.Error as exc:
                self.log_error("Could not save report: %s", exc)
                entries = self.faq()
                if entries is None:
                    return
                self.send_content(500, page(query, search_faq(query, entries), error="Could not save your report. Please try again.", details=details, mock_incident=mock_incident))
                return
            redirect_params = {"q": query, "submitted": "incident" if mock_incident else "report"}
            if mock_incident:
                redirect_params["incident"] = str(report_id)
            self.send_content(303, "", headers={"Location": f"/?{urlencode(redirect_params)}"})

    return Handler


def main():
    faq_path = Path(os.environ.get("TED_FAQ_PATH", str(FAQ_PATH)))
    load_faq(faq_path)
    database = Path(os.environ.get("TED_DB_PATH", str(DEFAULT_DB)))
    database.parent.mkdir(parents=True, exist_ok=True)
    initialize_db(database)
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer(("127.0.0.1", port), make_handler(faq_path, database))
    print(f"Ted Support Assistant running at http://127.0.0.1:{server.server_port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
