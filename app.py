import hmac
import io
import os
import re
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from functools import wraps
from pathlib import Path

from flask import (
    Flask,
    abort,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.environ.get("SECRET_KEY") or secrets.token_hex(32),
    DATABASE=str(Path(__file__).resolve().parent / "Register.db"),
    ADMIN_PASSWORD=os.environ.get("ADMIN_PASSWORD"),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
)

EMAIL_PATTERN = re.compile(
    r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$"
)
LEGACY_EXCEL_FILE = Path(__file__).resolve().parent / "Register.xlsx"
EXCEL_HEADERS = ["Бүртгэсэн огноо", "Овог", "Нэр", "Утасны дугаар", "Имэйл"]


def connect_db():
    connection = sqlite3.connect(app.config["DATABASE"])
    connection.row_factory = sqlite3.Row
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS registrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            surname TEXT NOT NULL,
            given_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            email TEXT NOT NULL
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS app_metadata (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    import_legacy_excel(connection)
    return connection


def import_legacy_excel(connection):
    connection.execute("BEGIN IMMEDIATE")
    try:
        imported = connection.execute(
            "SELECT 1 FROM app_metadata WHERE key = 'legacy_excel_imported'"
        ).fetchone()
        if imported:
            connection.commit()
            return

        if LEGACY_EXCEL_FILE.exists():
            workbook = load_workbook(LEGACY_EXCEL_FILE, read_only=True, data_only=True)
            worksheet = workbook.active
            headers = [cell.value for cell in worksheet[1]]
            if headers != EXCEL_HEADERS:
                raise ValueError(
                    f"Unexpected columns in legacy workbook: {LEGACY_EXCEL_FILE}"
                )

            for row_number, row in enumerate(
                worksheet.iter_rows(min_row=2, values_only=True), start=2
            ):
                if not any(value is not None for value in row):
                    continue
                if len(row) != len(EXCEL_HEADERS) or any(
                    value is None for value in row
                ):
                    raise ValueError(
                        f"Incomplete registration in legacy workbook row {row_number}"
                    )
                connection.execute(
                    """
                    INSERT INTO registrations
                        (created_at, surname, given_name, phone, email)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    tuple(str(value) for value in row),
                )
            workbook.close()

        connection.execute(
            "INSERT INTO app_metadata (key, value) VALUES ('legacy_excel_imported', '1')"
        )
        connection.commit()
    except Exception:
        connection.rollback()
        app.logger.exception("Failed to import the legacy Excel registrations")
        raise


@contextmanager
def database():
    connection = connect_db()
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login"))
        return view(*args, **kwargs)

    return wrapped


@app.route("/", methods=["GET", "POST"])
def register():
    form = {"surname": "", "given_name": "", "phone": "", "email": ""}
    message = None
    message_type = None
    status = 200

    if request.method == "POST":
        form = {
            "surname": request.form.get("surname", "").strip(),
            "given_name": request.form.get("given_name", "").strip(),
            "phone": request.form.get("phone", "").strip(),
            "email": request.form.get("email", "").strip(),
        }
        phone_digits = re.sub(r"[\s()+-]", "", form["phone"])
        errors = []

        if not form["surname"] or len(form["surname"]) > 100:
            errors.append("Овгоо оруулна уу (100 тэмдэгтээс хэтрэхгүй).")
        if not form["given_name"] or len(form["given_name"]) > 100:
            errors.append("Нэрээ оруулна уу (100 тэмдэгтээс хэтрэхгүй).")
        if not re.fullmatch(r"[0-9]{8}", phone_digits):
            errors.append("Утасны дугаарыг 8 оронтойгоор оруулна уу. Жишээ: 9911 2233.")
        if not EMAIL_PATTERN.fullmatch(form["email"]) or len(form["email"]) > 254:
            errors.append("Имэйл хаягаа зөв форматаар оруулна уу. Жишээ: ner@example.com.")

        if errors:
            message = " ".join(errors)
            message_type = "error"
            status = 400
        else:
            try:
                with database() as connection:
                    connection.execute(
                        """
                        INSERT INTO registrations
                            (created_at, surname, given_name, phone, email)
                        VALUES (?, ?, ?, ?, ?)
                        """,
                        (
                            datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S"),
                            form["surname"],
                            form["given_name"],
                            phone_digits,
                            form["email"],
                        ),
                    )
            except sqlite3.Error:
                app.logger.exception("Failed to save registration")
                abort(500)
            else:
                message = "Бүртгэл амжилттай хийгдлээ. Баярлалаа!"
                message_type = "success"
                form = {"surname": "", "given_name": "", "phone": "", "email": ""}

    return render_template(
        "index.html", form=form, message=message, message_type=message_type
    ), status


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if session.get("is_admin"):
        return redirect(url_for("admin_dashboard"))

    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)
    message = None

    if request.method == "POST":
        submitted_token = request.form.get("csrf_token", "")
        if not hmac.compare_digest(session["csrf_token"], submitted_token):
            abort(400)

        configured_password = app.config.get("ADMIN_PASSWORD")
        if not configured_password:
            message = "Админы нууц үг тохируулагдаагүй байна. README зааврыг дагана уу."
        elif hmac.compare_digest(
            configured_password, request.form.get("password", "")
        ):
            session.clear()
            session["is_admin"] = True
            return redirect(url_for("admin_dashboard"))
        else:
            message = "Нууц үг буруу байна."

    return render_template(
        "admin_login.html", csrf_token=session["csrf_token"], message=message
    )


@app.post("/admin/logout")
@admin_required
def admin_logout():
    submitted_token = request.form.get("csrf_token", "")
    if not hmac.compare_digest(session.get("csrf_token", ""), submitted_token):
        abort(400)
    session.clear()
    return redirect(url_for("admin_login"))


@app.get("/admin")
@admin_required
def admin_dashboard():
    with database() as connection:
        registrations = connection.execute(
            """
            SELECT id, created_at, surname, given_name, phone, email
            FROM registrations
            ORDER BY id DESC
            """
        ).fetchall()
    csrf_token = session.setdefault("csrf_token", secrets.token_urlsafe(32))
    return render_template(
        "admin.html", registrations=registrations, csrf_token=csrf_token
    )


@app.get("/admin/export.xlsx")
@admin_required
def export_excel():
    try:
        with database() as connection:
            registrations = connection.execute(
                """
                SELECT created_at, surname, given_name, phone, email
                FROM registrations
                ORDER BY id ASC
                """
            ).fetchall()
    except sqlite3.Error:
        app.logger.exception("Failed to read registrations for export")
        abort(500)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Бүртгэл"
    worksheet.append(["Бүртгэсэн огноо", "Овог", "Нэр", "Утасны дугаар", "Имэйл"])
    for registration in registrations:
        worksheet.append(
            [
                registration["created_at"],
                registration["surname"],
                registration["given_name"],
                registration["phone"],
                registration["email"],
            ]
        )
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="244B4A")
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.data_type = "s"
    worksheet.freeze_panes = "A2"
    for column, width in zip("ABCDE", (23, 20, 20, 20, 32)):
        worksheet.column_dimensions[column].width = width

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="Register.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
