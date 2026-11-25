import http.client
from contextlib import closing
from datetime import date, timedelta
from io import BytesIO
from pathlib import Path
import shutil
import sqlite3
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from urllib.parse import urlencode

from openpyxl import load_workbook
from pypdf import PdfWriter

import app


class WorkbookTests(unittest.TestCase):
    def test_search_sample_workbook(self):
        entries = app.load_faq(app.FAQ_PATH)
        results = app.search_faq("PDF/A filename", entries)
        self.assertEqual(results[0]["id"], "DEMO-002")
        self.assertIn("TED", results[0]["reference"])
        self.assertEqual(app.search_faq("platypus", entries), [])
        self.assertEqual(app.search_faq("what is the", entries), [])

    def test_bad_headers_and_duplicate_ids_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "faq.xlsx"
            shutil.copyfile(app.FAQ_PATH, path)
            workbook = load_workbook(path)
            workbook["FAQ"]["A1"] = "wrong"
            workbook.save(path)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "headers"):
                app.load_faq(path)
            workbook = load_workbook(path)
            workbook["FAQ"]["A1"] = "id"
            workbook["FAQ"]["A3"] = workbook["FAQ"]["A2"].value
            workbook.save(path)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "duplicate"):
                app.load_faq(path)

    def test_existing_mock_pbi_schema_migrates_without_losing_records(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "old.sqlite3"
            with closing(sqlite3.connect(path)) as connection, connection:
                connection.execute(app.MOCK_PBI_SCHEMA.replace("pdf_sha256 TEXT UNIQUE", "pdf_sha256 TEXT NOT NULL UNIQUE"))
                connection.execute(
                    """INSERT INTO mock_pbis (requester, operation, environment, filename, language,
                    document_type, product_type, effective_date, new_key, old_key, pdf_sha256)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    ("Demo", "add", "non-prod", "a.pdf", "English", "Demo", "Demo", "2099-01-01", "OLD", "", "hash"),
                )
            app.initialize_db(path)
            app.initialize_db(path)
            with closing(sqlite3.connect(path)) as connection:
                self.assertEqual(connection.execute("SELECT id, new_key, pdf_sha256 FROM mock_pbis").fetchall(), [(1, "OLD", "hash")])
                column = next(row for row in connection.execute("PRAGMA table_info(mock_pbis)") if row[1] == "pdf_sha256")
                self.assertEqual(column[3], 0)


class PortalTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.database = Path(cls.temp_dir.name) / "reports.sqlite3"
        cls.faq = Path(cls.temp_dir.name) / "faq.xlsx"
        shutil.copyfile(app.FAQ_PATH, cls.faq)
        app.initialize_db(cls.database)
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), app.make_handler(cls.faq, cls.database))
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join()
        cls.temp_dir.cleanup()

    def request(self, method, path, body=None, headers=None):
        connection = http.client.HTTPConnection("127.0.0.1", self.server.server_port, timeout=5)
        try:
            connection.request(method, path, body=body, headers=headers or {})
            response = connection.getresponse()
            return response.status, dict(response.getheaders()), response.read().decode("utf-8")
        finally:
            connection.close()

    def post(self, path, fields):
        return self.request("POST", path, urlencode(fields), {"Content-Type": "application/x-www-form-urlencoded"})

    def test_portal_identity_and_prototype_notices(self):
        for path in ("/", "/intake"):
            status, _, body = self.request("GET", path)
            self.assertEqual(status, 200)
            self.assertIn("TED Support", body)
            self.assertIn("RABOBANK", body)
            self.assertNotIn("hackathon", body.lower())
            self.assertIn("LOCAL PROTOTYPE", body)
        _, _, body = self.request("GET", "/")
        self.assertIn("UNOFFICIAL LOCAL PROTOTYPE", body)
        self.assertIn("File checks are local", body)

    def test_search_uses_workbook_and_reflects_saved_edits(self):
        status, _, body = self.request("GET", "/?q=PDF%2FA+filename")
        self.assertEqual(status, 200)
        self.assertIn("DEMO-002", body)
        self.assertIn("Demo placeholder - TED to provide approved source", body)
        workbook = load_workbook(self.faq)
        workbook["FAQ"]["D3"] = "Changed response for the test."
        workbook["FAQ"]["F3"] = "https://example.com/faq"
        workbook.save(self.faq)
        workbook.close()
        status, _, body = self.request("GET", "/?q=PDF%2FA+filename")
        self.assertEqual(status, 200)
        self.assertIn("Changed response for the test.", body)
        self.assertIn('href="https://example.com/faq"', body)

    def test_unmatched_search_and_escaped_input(self):
        status, _, body = self.request("GET", "/?q=%3Cscript%3Eplatypus%3C%2Fscript%3E")
        self.assertEqual(status, 200)
        self.assertIn("NO MATCHES FOUND", body)
        self.assertNotIn("<script>", body)
        self.assertIn("&lt;script&gt;", body)

    def test_bad_workbook_returns_an_error(self):
        try:
            self.faq.write_bytes(b"not an excel workbook")
            status, _, body = self.request("GET", "/?q=PDF")
            self.assertEqual(status, 503)
            self.assertIn("workbook is unavailable or invalid", body)
        finally:
            shutil.copyfile(app.FAQ_PATH, self.faq)

    def test_report_and_mock_incident_are_saved_locally(self):
        for mock in (False, True):
            fields = {"query": "retention", "details": "Where is retention documented?"}
            if mock:
                fields["mock_incident"] = "1"
            status, headers, _ = self.post("/report", fields)
            self.assertEqual(status, 303)
            status, _, body = self.request("GET", headers["Location"])
            self.assertEqual(status, 200)
            if mock:
                self.assertIn("mock Halo incident MOCK-", body)
                self.assertIn("Nothing was sent to Halo", body)
            else:
                self.assertIn("report was saved locally", body)
        with closing(sqlite3.connect(self.database)) as connection:
            rows = connection.execute(
                "SELECT query, details, mock_incident FROM reports WHERE query = ? ORDER BY id",
                ("retention",),
            ).fetchall()
        self.assertEqual(rows, [
            ("retention", "Where is retention documented?", 0),
            ("retention", "Where is retention documented?", 1),
        ])

    def test_invalid_report_preserves_input(self):
        status, _, body = self.post("/report", {"query": "retention", "details": " ", "mock_incident": "1"})
        self.assertEqual(status, 400)
        self.assertIn("Enter a search", body)
        self.assertIn('value="retention"', body)
        self.assertIn('value="1" checked', body)

    def test_valid_submission_creates_one_local_mock_pbi_and_rejects_duplicates(self):
        fields = self.intake_fields()
        status, _, body = self.post_documents(fields)
        self.assertEqual(status, 200)
        self.assertIn("Mock PBI DEMO-", body)
        self.assertIn("no real service was called", body)
        with closing(sqlite3.connect(self.database)) as connection:
            row = connection.execute("SELECT requester, filename, new_key, old_key FROM mock_pbis WHERE new_key = ?", (fields["new_key"],)).fetchone()
        self.assertEqual(row, ("Demo team", "test.pdf", "NEW-1", "OLD-1"))
        status, _, body = self.post_documents(fields)
        self.assertEqual(status, 400)
        self.assertIn("Duplicates are rejected", body)
        changed_key = dict(fields, new_key="NEW-2")
        status, _, body = self.post_documents(changed_key)
        self.assertEqual(status, 400)
        self.assertIn("Duplicates are rejected", body)

    def test_downloadable_samples_work_with_documented_fields(self):
        fields = self.intake_fields()
        fields.update(operation="add", filename="sample.pdf", new_key="DEMO-001", old_key="")
        status, _, body = self.post_documents(fields, workbook=app.sample_template(), pdf_name="sample.pdf", pdf_content=app.sample_pdf())
        self.assertEqual(status, 200)
        self.assertIn("Mock PBI DEMO-", body)

    def test_single_file_submissions_mark_unavailable_checks_as_skipped(self):
        fields = self.intake_fields()
        fields["new_key"] = "EXCEL-ONLY-1"
        status, _, body = self.post_documents(fields, include_pdf=False)
        self.assertEqual(status, 200)
        self.assertIn("Some checks were skipped", body)
        self.assertIn("PDF/A simulation was not run", body)
        self.assertIn("Duplicate PDF check was not performed", body)
        fields["new_key"] = "EXCEL-ONLY-2"
        status, _, _ = self.post_documents(fields, include_pdf=False)
        self.assertEqual(status, 200)
        fields["new_key"] = "PDF-ONLY-1"
        pdf = BytesIO()
        writer = PdfWriter()
        writer.add_blank_page(width=333, height=222)
        writer.write(pdf)
        status, _, body = self.post_documents(fields, include_workbook=False, pdf_content=pdf.getvalue())
        self.assertEqual(status, 200)
        self.assertIn("Excel template and spreadsheet checks were not performed", body)
        self.assertIn("PDF-to-Excel filename matching was not checked", body)
        fields["new_key"] = "PDF-ONLY-2"
        status, _, body = self.post_documents(fields, include_workbook=False, pdf_content=pdf.getvalue())
        self.assertEqual(status, 400)
        self.assertIn("Duplicates are rejected", body)
        status, _, body = self.post_documents(fields, include_workbook=False, include_pdf=False)
        self.assertEqual(status, 400)
        self.assertIn("Select at least one Excel workbook or PDF", body)

    def test_invalid_and_simulated_failed_checks_create_no_pbi(self):
        cases = [
            ({"old_key": ""}, "old document key is required"),
            ({"effective_date": "2026-02-30"}, "valid effective date"),
            ({"effective_date": (date.today() - timedelta(days=1)).isoformat()}, "cannot be before today"),
            ({"exception": "1"}, "exceptions are rejected"),
            ({"scenario": "pdfa"}, "PDF/A: simulated failure"),
            ({"scenario": "malware"}, "Malware: simulated failure"),
            ({"scenario": "dlp"}, "DLP: simulated failure"),
            ({"scenario": "permission"}, "Service-account permissions: simulated failure"),
        ]
        with closing(sqlite3.connect(self.database)) as connection:
            before = connection.execute("SELECT COUNT(*) FROM mock_pbis").fetchone()[0]
        for changes, expected in cases:
            with self.subTest(changes=changes):
                fields = self.intake_fields()
                fields.update(changes)
                status, _, body = self.post_documents(fields)
                self.assertEqual(status, 400)
                self.assertIn(expected, body)
        status, _, body = self.post_documents(self.intake_fields(), workbook=app.FAQ_PATH.read_bytes())
        self.assertEqual(status, 400)
        self.assertIn("use the demo template", body)
        status, _, body = self.post_documents(self.intake_fields(), pdf_name="wrong.pdf")
        self.assertEqual(status, 400)
        self.assertIn("Uploaded PDF filename must match", body)
        template = load_workbook(BytesIO(app.sample_template()))
        template["Document"]["A2"] = "OLD-VERSION"
        stream = BytesIO()
        template.save(stream)
        template.close()
        status, _, body = self.post_documents(self.intake_fields(), workbook=stream.getvalue())
        self.assertEqual(status, 400)
        self.assertIn("Template version must be", body)
        status, _, body = self.post_documents(self.intake_fields(), pdf_content=b"not a PDF")
        self.assertEqual(status, 400)
        self.assertIn("document could not be opened", body)
        with closing(sqlite3.connect(self.database)) as connection:
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM mock_pbis").fetchone()[0], before)

    def test_upload_limits_and_malformed_uploads(self):
        fields = self.intake_fields()
        status, _, body = self.post_documents(fields, workbook=b"x" * (app.MAX_FILE_SIZE + 1))
        self.assertEqual(status, 400)
        self.assertIn("Each file must be 5 MB or smaller", body)
        status, _, body = self.request("POST", "/intake", b"not multipart", {"Content-Type": "multipart/form-data; boundary=missing"})
        self.assertEqual(status, 400)
        self.assertIn("Invalid multipart upload", body)
        status, _, _ = self.post("/intake", fields)
        self.assertEqual(status, 415)

    def post_documents(self, fields, workbook=None, pdf_name="test.pdf", pdf_content=None, include_workbook=True, include_pdf=True):
        if workbook is None:
            template = load_workbook(BytesIO(app.sample_template()))
            for index, key in enumerate(app.DEMO_COLUMNS[1:], 2):
                template["Document"].cell(2, index, fields.get(key, ""))
            stream = BytesIO()
            template.save(stream)
            template.close()
            workbook = stream.getvalue()
        pdf = BytesIO()
        writer = PdfWriter()
        writer.add_blank_page(width=100, height=100)
        writer.write(pdf)
        boundary = "ted-test-boundary"
        parts = []
        for field, value in fields.items():
            parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="{field}"\r\n\r\n{value}\r\n'.encode())
        attachments = (
            ("workbook", "sample.xlsx" if include_workbook else "", workbook if include_workbook else b""),
            ("pdf", pdf_name if include_pdf else "", (pdf.getvalue() if pdf_content is None else pdf_content) if include_pdf else b""),
        )
        for field, filename, content in attachments:
            parts.append(
                f'--{boundary}\r\nContent-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'
                'Content-Type: application/octet-stream\r\n\r\n'.encode() + content + b"\r\n"
            )
        body = b"".join(parts) + f"--{boundary}--\r\n".encode()
        return self.request("POST", "/intake", body, {"Content-Type": f"multipart/form-data; boundary={boundary}"})

    @staticmethod
    def intake_fields():
        return {
            "requester": "Demo team", "operation": "replace", "environment": "non-prod",
            "filename": "test.pdf", "language": "English", "document_type": "Demo",
            "product_type": "Demo", "effective_date": date.today().isoformat(), "new_key": "NEW-1",
            "old_key": "OLD-1", "checklist": "1", "scenario": "pass",
        }


if __name__ == "__main__":
    unittest.main()
