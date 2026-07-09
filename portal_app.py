#!/usr/bin/env python3
"""
HR Employee Portal for iFocus Birthday Automation
-----------------------------------------------------
A small LAN-only web app so HR can manage the employee list (name, email,
dob) that drives send_birthday_emails.py, without touching Excel files or
code directly.

Features:
  - Single shared-password login (PORTAL_PASSWORD in .env)
  - View / add / edit / delete employees
  - Bulk upload an .xlsx to import/update many employees at once
  - Export the current list back to .xlsx (backup/download)

Run for local development:
    python portal_app.py

For always-on hosting on a server/NAS, use run_portal.py instead (see
README.md).
"""

import functools
import os
import secrets
import tempfile

from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

import employee_store
from employee_store import ValidationError

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

PORTAL_PASSWORD = os.environ.get("PORTAL_PASSWORD")
SECRET_KEY = os.environ.get("SECRET_KEY")

app = Flask(__name__)
app.config["SECRET_KEY"] = SECRET_KEY or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024  # 5 MB upload cap

if not SECRET_KEY:
    print(
        "WARNING: SECRET_KEY is not set in .env - using a random key that will "
        "change every restart (this logs everyone out on each restart). "
        "Add SECRET_KEY to .env for a stable session key."
    )


def login_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)

    return wrapped


@app.route("/login", methods=["GET", "POST"])
def login():
    if not PORTAL_PASSWORD:
        return (
            "PORTAL_PASSWORD is not set on the server. Ask IT to add it to .env "
            "before the portal can be used.",
            500,
        )

    if request.method == "POST":
        entered = request.form.get("password", "")
        if secrets.compare_digest(entered, PORTAL_PASSWORD):
            session.clear()
            session["logged_in"] = True
            next_url = request.args.get("next") or url_for("dashboard")
            return redirect(next_url)
        flash("Incorrect password.", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    employees = employee_store.list_employees()
    return render_template("dashboard.html", employees=employees)


@app.route("/employees", methods=["POST"])
@login_required
def add_employee():
    try:
        employee_store.add_employee(
            request.form.get("name", ""),
            request.form.get("email", ""),
            request.form.get("dob", ""),
        )
        flash("Employee added.", "success")
    except ValidationError as e:
        flash(str(e), "error")
    return redirect(url_for("dashboard"))


@app.route("/employees/<int:emp_id>/edit", methods=["POST"])
@login_required
def edit_employee(emp_id):
    try:
        employee_store.update_employee(
            emp_id,
            request.form.get("name", ""),
            request.form.get("email", ""),
            request.form.get("dob", ""),
        )
        flash("Employee updated.", "success")
    except ValidationError as e:
        flash(str(e), "error")
    return redirect(url_for("dashboard"))


@app.route("/employees/<int:emp_id>/delete", methods=["POST"])
@login_required
def delete_employee(emp_id):
    try:
        employee_store.delete_employee(emp_id)
        flash("Employee deleted.", "success")
    except ValidationError as e:
        flash(str(e), "error")
    return redirect(url_for("dashboard"))


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    file = request.files.get("file")
    if not file or not file.filename:
        flash("Choose an .xlsx file to upload first.", "error")
        return redirect(url_for("dashboard"))

    if not file.filename.lower().endswith(".xlsx"):
        flash("Only .xlsx files are supported.", "error")
        return redirect(url_for("dashboard"))

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    try:
        os.close(tmp_fd)
        file.save(tmp_path)
        try:
            summary = employee_store.import_from_xlsx(tmp_path)
        except ValidationError as e:
            flash(str(e), "error")
            return redirect(url_for("dashboard"))
    finally:
        os.remove(tmp_path)

    msg = f"Import complete: {summary['added']} added, {summary['updated']} updated"
    if summary["skipped"]:
        msg += f", {len(summary['skipped'])} skipped"
    flash(msg, "success")

    for row_num, reason in summary["skipped"][:10]:
        flash(f"Row {row_num} skipped: {reason}", "warning")

    return redirect(url_for("dashboard"))


@app.route("/export")
@login_required
def export():
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)
    employee_store.export_to_xlsx(tmp_path)
    return send_file(
        tmp_path,
        as_attachment=True,
        download_name="employees_export.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


employee_store.init_db()

if __name__ == "__main__":
    # Dev server only. For always-on LAN hosting, use run_portal.py.
    app.run(host="0.0.0.0", port=5000, debug=False)
