#!/usr/bin/env python3
"""
Automated Birthday Email Sender for iFocus Systec
---------------------------------------------------
Reads employee data from employees.db (kept up to date via the HR Portal,
see portal_app.py), finds anyone whose birthday (month + day) matches today,
and sends them the HR birthday email template via SMTP - with the same
cupcake image and BCC-to-allemployees behavior seen in the existing manual
emails.

USAGE:
    python3 send_birthday_emails.py
    python3 send_birthday_emails.py --dry-run
    python3 send_birthday_emails.py --dry-run --test-date 2026-07-06
    python3 send_birthday_emails.py --loop              # run forever, checks daily at 09:00
    python3 send_birthday_emails.py --loop --loop-hour 8 # run forever, checks daily at 08:00

Either run once via a cron job/Task Scheduler (see README.md), or start it
once with --loop and leave the process running - it checks for birthdays
once a day and sleeps in between.
"""

import argparse
import csv
import logging
import os
import smtplib
import ssl
import time
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from email.utils import formataddr

from dotenv import load_dotenv

import employee_store

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

load_dotenv(os.path.join(BASE_DIR, ".env"))

SMTP_HOST = os.environ.get("SMTP_HOST", "mail.ifocussystec.in")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))               # 465 = implicit TLS/SSL
SMTP_SECURE = os.environ.get("SMTP_SECURE", "true").strip().lower() not in ("false", "0", "")
SMTP_USER = os.environ.get("SMTP_USER", "itsupport@ifocussystec.com")
# SMTP_PASSWORD is the current name; SMTP_PASS kept as a fallback for old .env files.
SMTP_PASS = os.environ.get("SMTP_PASSWORD") or os.environ.get("SMTP_PASS")

FROM_NAME = os.environ.get("SMTP_FROM_NAME", "HR Support")
BCC_ADDRESS = "ramkirangaruda2006@gmail.com"  # TODO: change back to allemployees@ifocussystec.com after testing

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
    parser.add_argument(
        "--loop",
        action="store_true",
        help="Run forever: check for birthdays once a day instead of exiting "
        "after one run. Use this instead of a cron/Task Scheduler job - just "
        "start it once (e.g. in the background) and leave it running.",
    )
    parser.add_argument(
        "--loop-hour",
        type=int,
        default=9,
        metavar="HOUR",
        help="Hour of the day (0-23, local time) to send at when --loop is used. Default: 9.",
    )
    return parser.parse_args()


def resolve_today(test_date_str):
    if not test_date_str:
        return date.today()
    try:
        return date.fromisoformat(test_date_str)
    except ValueError:
        raise SystemExit(f"--test-date must be in YYYY-MM-DD format, got: {test_date_str!r}")


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
            if SMTP_SECURE or SMTP_PORT == 465:
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


def check_and_send(today, dry_run):
    """Look up today's birthdays and send (or dry-run save) each one."""
    employee_store.init_db()
    logger.info("Reading employee data from %s", employee_store.DB_PATH)
    birthdays = employee_store.get_todays_birthdays(today)
    if not birthdays:
        logger.info("No birthdays today (%s).", today.isoformat())
        return

    for name, email in birthdays:
        if already_sent_today(name, today):
            logger.info("Already sent to %s today, skipping.", name)
            continue

        msg = build_message(name, email)

        if dry_run:
            save_dry_run(msg, name, today)
            continue

        try:
            send_email(msg)
            log_sent(name, email, today)
            logger.info("Sent birthday email to %s <%s>", name, email)
        except Exception as e:
            logger.error("FAILED to send to %s <%s>: %s", name, email, e)


def seconds_until_next_run(loop_hour, now=None):
    now = now or datetime.now()
    target = now.replace(hour=loop_hour, minute=0, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)
    return (target - now).total_seconds()


def run_loop(loop_hour, dry_run):
    logger.info(
        "Starting in --loop mode: will check for birthdays once a day at %02d:00. "
        "Leave this process running instead of using cron/Task Scheduler.",
        loop_hour,
    )
    while True:
        try:
            check_and_send(date.today(), dry_run)
        except Exception:
            logger.exception("Unexpected error during daily birthday check - will retry tomorrow.")

        wait_seconds = seconds_until_next_run(loop_hour)
        logger.info("Next check in %.1f hours (around %02d:00).", wait_seconds / 3600, loop_hour)
        time.sleep(wait_seconds)


def main():
    args = parse_args()
    setup_logging()

    if args.dry_run:
        logger.info("Running in --dry-run mode (no emails will be sent).")

    if not args.dry_run and not SMTP_PASS:
        raise SystemExit("SMTP_PASSWORD is not set. Add it to .env before running (see .env.example).")

    if args.loop:
        if not 0 <= args.loop_hour <= 23:
            raise SystemExit("--loop-hour must be between 0 and 23.")
        run_loop(args.loop_hour, args.dry_run)
        return

    today = resolve_today(args.test_date)
    if args.test_date:
        logger.info("Simulating today as %s.", today.isoformat())
    check_and_send(today, args.dry_run)


if __name__ == "__main__":
    main()
