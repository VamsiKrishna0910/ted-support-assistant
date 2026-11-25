import http.client
from contextlib import closing
from pathlib import Path
import shutil
import sqlite3
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from urllib.parse import urlencode

from openpyxl import load_workbook

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

    def test_guided_intake_rejects_missing_old_key_and_checklist(self):
        with closing(sqlite3.connect(self.database)) as connection:
            before = connection.execute("SELECT COUNT(*) FROM document_requests").fetchone()[0]
        fields = self.intake_fields()
        fields["old_key"] = ""
        fields.pop("checklist")
        status, _, body = self.post("/intake", fields)
        self.assertEqual(status, 400)
        self.assertIn("old document key is required", body)
        self.assertIn("Review and confirm the checklist", body)
        self.assertIn('value="test.pdf"', body)
        with closing(sqlite3.connect(self.database)) as connection:
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM document_requests").fetchone()[0], before)

    def test_guided_intake_rejects_bad_date_and_saves_valid_request(self):
        fields = self.intake_fields()
        fields["effective_date"] = "2026-02-30"
        status, _, body = self.post("/intake", fields)
        self.assertEqual(status, 400)
        self.assertIn("valid effective date", body)
        fields["effective_date"] = "2026-12-01"
        status, headers, _ = self.post("/intake", fields)
        self.assertEqual(status, 303)
        status, _, body = self.request("GET", headers["Location"])
        self.assertEqual(status, 200)
        self.assertIn("not sent to TED", body)
        with closing(sqlite3.connect(self.database)) as connection:
            row = connection.execute(
                "SELECT operation, filename, new_key, old_key, checklist_confirmed FROM document_requests"
            ).fetchone()
        self.assertEqual(row, ("replace", "test.pdf", "NEW-1", "OLD-1", 1))

    @staticmethod
    def intake_fields():
        return {
            "requester": "Demo team", "operation": "replace", "environment": "non-prod",
            "filename": "test.pdf", "language": "English", "document_type": "Demo",
            "product_type": "Demo", "effective_date": "2026-12-01", "new_key": "NEW-1",
            "old_key": "OLD-1", "checklist": "1",
        }


if __name__ == "__main__":
    unittest.main()
