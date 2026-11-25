"""Local TED support pilot with an Excel FAQ and guided intake."""

from datetime import date, datetime
from contextlib import closing
import html
import os
from pathlib import Path
import re
import sqlite3
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlencode, urlsplit
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


ROOT = Path(__file__).resolve().parent
FAQ_PATH = ROOT / "data" / "ted_faq.xlsx"
DEFAULT_DB = ROOT / "data" / "reports.sqlite3"
MAX_BODY = 16_384
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


def initialize_db(path):
    with closing(sqlite3.connect(path)) as connection, connection:
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


def header():
    return """<header class="topbar"><a class="brand" href="/"><span class="brand-mark">t<span>✳</span></span><span>ted<span class="brand-divider"></span><small>SUPPORT ASSISTANT</small></span></a><span class="demo-badge"><span></span> LOCAL PILOT</span></header>"""


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
            f'<h2>Answers from the workbook</h2><p>Showing {len(results)} matching FAQ entries.</p>'
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
<main><div class="hero"><div class="hero-copy"><div class="pill"><span>✦</span> TED SUPPORT PILOT</div>
<h1>Good answers.<br><em>Less digging.</em></h1>
<p>Search common TED questions or prepare a generic-document request.</p>
<form class="search" action="/" method="get"><span class="search-icon" aria-hidden="true">⌕</span>
<label class="sr-only" for="query">Search the FAQ</label><input id="query" name="q" value="{escape(query, quote=True)}" placeholder="Ask about documents or onboarding..." maxlength="500" required>
<button type="submit">Search <span aria-hidden="true">→</span></button></form>
<div class="suggestions">TRY ASKING <a href="/?q=PDF%2FA">PDF/A</a><a href="/?q=document+keys">Document keys</a><a href="/?q=API+onboarding">API onboarding</a></div>
<a class="intake-link" href="/intake">Start a generic-document request →</a>
</div><div class="hero-art" aria-hidden="true"><div class="art-orbit orbit-one"></div><div class="art-orbit orbit-two"></div><div class="art-card card-back"><i></i><i></i><i></i></div><div class="art-card card-front"><span class="art-star">✳</span><div><i></i><i></i><i></i></div></div><span class="spark spark-one">✦</span><span class="spark spark-two">✳</span><span class="art-label">A LITTLE CLARITY, DELIVERED.</span></div></div>
<div class="pilot-banner"><strong>Sample content only.</strong> FAQ answers and intake checks require TED review; do not use them as production policy. No files are uploaded or verified.</div>
<div class="content"><section class="answers" aria-live="polite">{result_html}</section>
<aside class="report"><div class="report-icon">✳</div><span class="eyebrow">HELP US IMPROVE</span><h2>Something missing?</h2>
<p>Record a gap locally. Do not include secrets or customer data.</p>{alert(notice, "success")}{alert(error)}
<form method="post" action="/report"><label for="report-query">Your search</label><input id="report-query" name="query" value="{escape(query, quote=True)}" placeholder="e.g. document keys" maxlength="500" required>
<label for="details">What’s missing?</label><textarea id="details" name="details" placeholder="Tell us what you hoped to find..." maxlength="3000" required>{escape(details)}</textarea>
<label class="checkbox"><input type="checkbox" name="mock_incident" value="1"{checked}><span><strong>Create a mock Halo incident</strong><small>Demo only. Nothing is sent to Halo.</small></span></label>
<button class="report-button" type="submit">Report missing documentation <span aria-hidden="true">→</span></button></form>
<div class="report-foot">Saved locally on this computer; not sent to TED.</div></aside></div></main>
<footer><span><strong>ted</strong> &nbsp; / &nbsp; SUPPORT ASSISTANT</span><span>LOCAL PILOT · SAMPLE CONTENT</span></footer>
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


def intake_page(values=None, error="", notice=""):
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
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Generic-document intake | Ted Support Assistant</title>
<link rel="stylesheet" href="/style.css"></head><body>{header()}
<main class="source"><a class="back" href="/">← Back to search</a><span class="eyebrow">GUIDED PILOT INTAKE</span>
<h1>Generic-document request</h1>
<p class="intro">This sample form saves details locally only. It does not submit work to TED, upload files, check PDF/A, or validate keys. Do not enter real customer data or secrets.</p>
{alert(notice, "success")}{alert(error)}
<form class="intake-form" method="post" action="/intake">{operation}{environment}{"".join(inputs)}
<fieldset><legend>Before submitting</legend><p>Manually check the current template, PDF/A, matching filenames, accessible files, key uniqueness and effective date with TED’s approved process.</p>
<label class="checkbox"><input type="checkbox" name="checklist" value="1"{checked} required><span>I have reviewed this checklist. These checks are not automated.</span></label></fieldset>
<button class="report-button" type="submit">Save request locally →</button></form></main></body></html>"""


def validate_intake(form):
    values = {name: form.get(name, [""])[0].strip() for name, _, _ in INTAKE_FIELDS}
    for name in ("operation", "environment", "checklist"):
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
    if values["environment"] not in ("prod", "non-prod"):
        errors.append("Choose an environment.")
    if values["checklist"] != "1":
        errors.append("Review and confirm the checklist.")
    if values["effective_date"]:
        try:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", values["effective_date"]):
                raise ValueError("Invalid date format")
            date.fromisoformat(values["effective_date"])
        except ValueError:
            errors.append("Enter a valid effective date (YYYY-MM-DD).")
    return values, errors


def make_handler(faq_path, database):
    class Handler(BaseHTTPRequestHandler):
        def send_content(self, status, content, content_type="text/html; charset=utf-8", headers=None):
            body = content.encode("utf-8")
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
                notice = "Your request was saved locally only. It was not sent to TED." if params.get("submitted") == ["1"] else ""
                self.send_content(200, intake_page(notice=notice))
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
            if length < 0 or length > MAX_BODY:
                self.send_content(413, "Request body too large", "text/plain; charset=utf-8")
                return
            if self.headers.get_content_type() != "application/x-www-form-urlencoded":
                self.send_content(415, "Unsupported content type", "text/plain; charset=utf-8")
                return
            try:
                form = parse_qs(self.rfile.read(length).decode("utf-8"), keep_blank_values=True, max_num_fields=20)
            except (UnicodeDecodeError, ValueError):
                self.send_content(400, "Invalid form data", "text/plain; charset=utf-8")
                return
            if route == "/intake":
                self.save_intake(form)
            else:
                self.save_report(form)

        def save_intake(self, form):
            values, errors = validate_intake(form)
            if errors:
                self.send_content(400, intake_page(values, " ".join(errors)))
                return
            try:
                with closing(sqlite3.connect(database)) as connection, connection:
                    connection.execute(
                        """INSERT INTO document_requests
                        (requester, operation, environment, filename, language, document_type,
                         product_type, effective_date, new_key, old_key, checklist_confirmed)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (values["requester"], values["operation"], values["environment"], values["filename"],
                         values["language"], values["document_type"], values["product_type"],
                         values["effective_date"], values["new_key"], values["old_key"], 1),
                    )
            except sqlite3.Error as exc:
                self.log_error("Could not save request: %s", exc)
                self.send_content(500, intake_page(values, "Could not save your request. Please try again."))
                return
            self.send_content(303, "", headers={"Location": "/intake?submitted=1"})

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
