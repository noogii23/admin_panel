import io
import re
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

import app as registration_app


class RegistrationTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.legacy_excel_file = registration_app.LEGACY_EXCEL_FILE
        registration_app.app.config.update(
            DATABASE=str(Path(self.temp_dir.name) / "Register.db"),
            ADMIN_PASSWORD="test-admin-password",
            TESTING=True,
        )
        registration_app.LEGACY_EXCEL_FILE = Path(self.temp_dir.name) / "Register.xlsx"
        self.client = registration_app.app.test_client()

    def tearDown(self):
        registration_app.LEGACY_EXCEL_FILE = self.legacy_excel_file
        self.temp_dir.cleanup()

    def register(self, **overrides):
        data = {
            "surname": "Бат",
            "given_name": "Номин",
            "phone": "9911 2233",
            "email": "ner@example.com",
        }
        data.update(overrides)
        return self.client.post("/", data=data)

    def login(self, password="test-admin-password"):
        response = self.client.get("/admin/login")
        csrf_token = re.search(
            rb'name="csrf_token" value="([^"]+)"', response.data
        ).group(1).decode()
        return self.client.post(
            "/admin/login", data={"csrf_token": csrf_token, "password": password}
        )

    def test_registration_is_saved_to_sqlite(self):
        response = self.register(surname="=1+1")
        self.assertEqual(response.status_code, 200)
        self.assertIn("Бүртгэл амжилттай хийгдлээ".encode(), response.data)

        with registration_app.database() as connection:
            row = connection.execute(
                "SELECT surname, phone, email FROM registrations"
            ).fetchone()
        self.assertEqual(tuple(row), ("=1+1", "99112233", "ner@example.com"))

    def test_existing_excel_registrations_are_imported_once(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(registration_app.EXCEL_HEADERS)
        worksheet.append(
            ["2026-11-13 21:20:48", "Хуучин", "Бүртгэл", "99112233", "old@example.com"]
        )
        workbook.save(registration_app.LEGACY_EXCEL_FILE)

        with registration_app.database() as connection:
            row = connection.execute(
                "SELECT surname, email FROM registrations"
            ).fetchone()
        self.assertEqual(tuple(row), ("Хуучин", "old@example.com"))

        with registration_app.database() as connection:
            count = connection.execute(
                "SELECT COUNT(*) FROM registrations"
            ).fetchone()[0]
        self.assertEqual(count, 1)

    def test_invalid_email_is_rejected_without_database_record(self):
        response = self.register(email="ner@")
        self.assertEqual(response.status_code, 400)
        self.assertIn("ner@example.com".encode(), response.data)
        with registration_app.database() as connection:
            count = connection.execute(
                "SELECT COUNT(*) FROM registrations"
            ).fetchone()[0]
        self.assertEqual(count, 0)

    def test_invalid_phone_is_rejected_without_database_record(self):
        response = self.register(phone="9911")
        self.assertEqual(response.status_code, 400)
        with registration_app.database() as connection:
            count = connection.execute(
                "SELECT COUNT(*) FROM registrations"
            ).fetchone()[0]
        self.assertEqual(count, 0)

    def test_admin_pages_are_protected_and_login_shows_registrations(self):
        self.register()
        response = self.client.get("/admin")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/admin/login", response.headers["Location"])
        response = self.login("wrong-password")
        self.assertEqual(response.status_code, 200)
        self.assertIn("Нууц үг буруу байна".encode(), response.data)
        self.assertEqual(self.login().status_code, 302)
        response = self.client.get("/admin")
        self.assertEqual(response.status_code, 200)
        self.assertIn("ner@example.com".encode(), response.data)

    def test_excel_export_requires_login_and_contains_database_records(self):
        self.register(surname="=1+1")
        response = self.client.get("/admin/export.xlsx")
        self.assertEqual(response.status_code, 302)

        self.login()
        response = self.client.get("/admin/export.xlsx")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.mimetype,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        workbook = load_workbook(io.BytesIO(response.data))
        worksheet = workbook.active
        self.assertEqual(worksheet.max_row, 2)
        self.assertEqual(worksheet.cell(2, 2).value, "=1+1")
        self.assertEqual(worksheet.cell(2, 2).data_type, "s")


if __name__ == "__main__":
    unittest.main()
