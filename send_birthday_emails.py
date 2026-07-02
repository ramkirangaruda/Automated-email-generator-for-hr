#!/usr/bin/env python3
"""
Automated Birthday Email Sender for iFocus Systec
---------------------------------------------------
Reads employee data from an .xlsx file (name, email, dob columns), finds
anyone whose birthday (month + day) matches today, and sends them the HR
birthday email template via SMTP - with the same cupcake image and
BCC-to-allemployees behavior seen in the existing manual emails.

USAGE:
    python3 send_birthday_emails.py
    python3 send_birthday_emails.py --dry-run
    python3 send_birthday_emails.py --dry-run --test-date 2026-07-06

Meant to be triggered once a day by a cron job (see README.md for setup).
"""

import argparse
import csv
import logging
import os
import smtplib
import ssl
import time
from datetime import date
from email.message import EmailMessage
from email.utils import formataddr

import openpyxl
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

load_dotenv(os.path.join(BASE_DIR, ".env"))

SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")         # Gmail SMTP
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))               # 587 = STARTTLS for Gmail
SMTP_USER = os.environ.get("SMTP_USER", "hr.support@ifocussystec.com")
SMTP_PASS = os.environ.get("SMTP_PASS")  # set via .env - never hardcode

FROM_NAME = "HR Support"
BCC_ADDRESS = "allemployees@ifocussystec.com"

# Real employee data belongs in employees.local.xlsx (gitignored, contains PII).
# employees.xlsx is the tracked, dummy/example file showing the expected format.
# Set EMPLOYEES_XLSX to point at a different path if needed.
_LOCAL_DATA_PATH = os.path.join(BASE_DIR, "employees.local.xlsx")
_EXAMPLE_DATA_PATH = os.path.join(BASE_DIR, "employees.xlsx")
EMPLOYEES_PATH = os.environ.get(
    "EMPLOYEES_XLSX",
    _LOCAL_DATA_PATH if os.path.exists(_LOCAL_DATA_PATH) else _EXAMPLE_DATA_PATH,
)

LOG_PATH = os.path.join(BASE_DIR, "sent_log.csv")
RUN_LOG_PATH = os.path.join(BASE_DIR, "birthday_automation.log")
IMAGE_PATH = os.path.join(BASE_DIR, "360_F_294637909_957UbRCZ8umRl6c6YzAcR78nAakfgSxf.jpg")
DRY_RUN_DIR = os.path.join(BASE_DIR, "dry_run_output")

SEND_RETRIES = 2          # number of retries after the first attempt
SEND_RETRY_BACKOFF = 5    # seconds, doubles after each retry

COMPANY_SIGNATURE_HTML = """
<p>Regards,</p>
<p><b>Team HR</b></p>
<p><b>iFocus Systec India Pvt Ltd</b></p>
<p>
  <a href="mailto:hr.support@ifocussystec.com">hr.support@ifocussystec.com</a>
  &nbsp;|&nbsp; Contact: +91-7899826062 / 9019702443
</p>
<p>www.ifocussystec.com</p>
<p><b>An ISO 9001:2015 Organization</b></p>
"""

HTML_TEMPLATE = """\
<html>
  <body style="font-family: Arial, sans-serif; color: #111;">
    <div style="text-align:left;">
      <img src="cid:birthday_img" alt="Happy Birthday" style="max-width:500px;" />
    </div>
    <p style="color:#d63384; font-weight:bold; font-size:16px;">{name}</p>
    <p>Wishing you joy today and success all year ahead. Your hard work is
    appreciated&mdash;may this year bring growth, new opportunities, and continued shine!</p>
    {signature}
  </body>
</html>
"""

logger = logging.getLogger("birthday_automation")


def setup_logging():
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    file_handler = logging.FileHandler(RUN_LOG_PATH, encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    )
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter("%(message)s"))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)


def parse_args():
    parser = argparse.ArgumentParser(description="Send HR birthday emails.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Build the email(s) and save as .eml files instead of sending via SMTP.",
    )
    parser.add_argument(
        "--test-date",
        metavar="YYYY-MM-DD",
        help="Pretend today is this date, instead of using the real current date.",
    )
    return parser.parse_args()


def resolve_today(test_date_str):
    if not test_date_str:
        return date.today()
    try:
        return date.fromisoformat(test_date_str)
    except ValueError:
        raise SystemExit(f"--test-date must be in YYYY-MM-DD format, got: {test_date_str!r}")


def load_todays_birthdays(xlsx_path, today):
    """Return list of (name, email) for people whose month/day matches today.

    Skips rows with missing/placeholder emails or malformed dates, logging
    each skip so nothing fails silently.
    """
    if not os.path.exists(xlsx_path):
        raise SystemExit(f"Employee data file not found: {xlsx_path}")

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        logger.warning("Employee data file %s is empty.", xlsx_path)
        return []

    header = [str(h).strip().lower() if h is not None else "" for h in header]
    try:
        name_idx = header.index("name")
        email_idx = header.index("email")
        dob_idx = header.index("dob")
    except ValueError:
        raise SystemExit(
            f"Employee data file {xlsx_path} must have 'name', 'email', 'dob' columns; "
            f"found: {header}"
        )

    matches = []
    for row_num, row in enumerate(rows, start=2):
        if row is None or all(cell is None for cell in row):
            continue

        name = str(row[name_idx]).strip() if row[name_idx] is not None else ""
        email = str(row[email_idx]).strip() if row[email_idx] is not None else ""
        dob_raw = row[dob_idx]

        if not name:
            logger.warning("Row %d skipped: missing name.", row_num)
            continue

        if not email or "TODO" in email.upper():
            logger.warning("Row %d (%s) skipped: missing/placeholder email (%r).", row_num, name, email)
            continue

        dob_date = _parse_dob(dob_raw)
        if dob_date is None:
            logger.warning("Row %d (%s) skipped: malformed date (%r).", row_num, name, dob_raw)
            continue

        if dob_date.month == today.month and dob_date.day == today.day:
            matches.append((name, email))

    wb.close()
    return matches


def _parse_dob(dob_raw):
    """Accept either a datetime/date object (openpyxl-parsed) or an ISO string."""
    if dob_raw is None:
        return None
    if isinstance(dob_raw, date):
        return dob_raw
    if hasattr(dob_raw, "date"):  # datetime
        return dob_raw.date()
    try:
        return date.fromisoformat(str(dob_raw).strip())
    except ValueError:
        return None


def already_sent_today(name, today):
    """Check sent_log.csv to avoid double-sending if the cron somehow runs twice."""
    if not os.path.exists(LOG_PATH):
        return False
    today_str = today.isoformat()
    with open(LOG_PATH, newline="", encoding="utf-8") as f:
        for row in csv.reader(f):
            if len(row) >= 2 and row[0] == today_str and row[1] == name:
                return True
    return False


def log_sent(name, email, today):
    file_exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["date", "name", "email"])
        writer.writerow([today.isoformat(), name, email])


def build_message(name, email):
    msg = EmailMessage()
    msg["Subject"] = f"Happy Birthday << {name} >>"
    msg["From"] = formataddr((FROM_NAME, SMTP_USER))
    msg["To"] = email
    msg["Bcc"] = BCC_ADDRESS

    html_body = HTML_TEMPLATE.format(name=name, signature=COMPANY_SIGNATURE_HTML)
    msg.set_content("Happy Birthday! (View this email in HTML to see the card.)")
    msg.add_alternative(html_body, subtype="html")

    # Embed the cupcake image inline so it always renders (no "remote resources blocked" issue)
    if os.path.exists(IMAGE_PATH):
        with open(IMAGE_PATH, "rb") as img:
            img_data = img.read()
        msg.get_payload()[1].add_related(
            img_data, maintype="image", subtype="jpeg", cid="<birthday_img>"
        )
    else:
        logger.warning("%s not found - email for %s will send without the image.", IMAGE_PATH, name)

    return msg


def save_dry_run(msg, name, today):
    os.makedirs(DRY_RUN_DIR, exist_ok=True)
    safe_name = "".join(c if c.isalnum() else "_" for c in name)
    out_path = os.path.join(DRY_RUN_DIR, f"{today.isoformat()}_{safe_name}.eml")
    with open(out_path, "wb") as f:
        f.write(bytes(msg))
    logger.info("[DRY RUN] Saved preview for %s -> %s", name, out_path)


def send_email(msg):
    """Send msg over SMTP, retrying on transient connection errors."""
    context = ssl.create_default_context()
    attempt = 0
    delay = SEND_RETRY_BACKOFF
    while True:
        try:
            if SMTP_PORT == 465:
                with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context, timeout=30) as server:
                    server.login(SMTP_USER, SMTP_PASS)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
                    server.starttls(context=context)
                    server.login(SMTP_USER, SMTP_PASS)
                    server.send_message(msg)
            return
        except (smtplib.SMTPException, OSError) as e:
            if attempt >= SEND_RETRIES:
                raise
            attempt += 1
            logger.warning(
                "SMTP send failed (attempt %d/%d): %s - retrying in %ds",
                attempt, SEND_RETRIES, e, delay,
            )
            time.sleep(delay)
            delay *= 2


def main():
    args = parse_args()
    setup_logging()

    today = resolve_today(args.test_date)

    if args.dry_run:
        logger.info("Running in --dry-run mode (no emails will be sent).")
    if args.test_date:
        logger.info("Simulating today as %s.", today.isoformat())

    if not args.dry_run and not SMTP_PASS:
        raise SystemExit("SMTP_PASS is not set. Add it to .env before running (see .env.example).")

    logger.info("Reading employee data from %s", EMPLOYEES_PATH)
    birthdays = load_todays_birthdays(EMPLOYEES_PATH, today)
    if not birthdays:
        logger.info("No birthdays today (%s).", today.isoformat())
        return

    for name, email in birthdays:
        if already_sent_today(name, today):
            logger.info("Already sent to %s today, skipping.", name)
            continue

        msg = build_message(name, email)

        if args.dry_run:
            save_dry_run(msg, name, today)
            continue

        try:
            send_email(msg)
            log_sent(name, email, today)
            logger.info("Sent birthday email to %s <%s>", name, email)
        except Exception as e:
            logger.error("FAILED to send to %s <%s>: %s", name, email, e)


if __name__ == "__main__":
    main()
