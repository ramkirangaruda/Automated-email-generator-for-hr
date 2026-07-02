#!/usr/bin/env python3
"""
Automated Birthday Email Sender for iFocus Systec
---------------------------------------------------
Reads employees.csv, finds anyone whose birthday (month + day) is today,
and sends them the HR birthday email template via SMTP - with the same
cupcake image and BCC-to-allemployees behavior seen in the existing manual emails.

USAGE:
    python3 send_birthday_emails.py

Meant to be triggered once a day by a cron job (see README.md for cPanel setup).
"""

import csv
import os
import smtplib
import ssl
from datetime import date
from email.message import EmailMessage
from email.utils import formataddr

# ---------------------------------------------------------------------------
# CONFIG - fill these in via environment variables (recommended) or here.
# NEVER commit real passwords into this file if it goes into git.
# ---------------------------------------------------------------------------

SMTP_HOST = os.environ.get("SMTP_HOST", "mail.ifocussystec.in")   # confirm exact host - see README
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))               # 465 = SSL, 587 = STARTTLS
SMTP_USER = os.environ.get("SMTP_USER", "hr.support@ifocussystec.com")
SMTP_PASS = os.environ.get("SMTP_PASS")  # set this as an env var / cPanel cron env - do not hardcode

FROM_NAME = "HR Support"
BCC_ADDRESS = "allemployees@ifocussystec.com"

CSV_PATH = os.path.join(os.path.dirname(__file__), "employees.csv")
LOG_PATH = os.path.join(os.path.dirname(__file__), "sent_log.csv")
IMAGE_PATH = os.path.join(os.path.dirname(__file__), "birthday_image.png")  # the cupcake graphic

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


def load_todays_birthdays(csv_path):
    """Return list of (name, email) for people whose month/day matches today."""
    today = date.today()
    matches = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            dob = row["dob"].strip()
            try:
                dob_date = date.fromisoformat(dob)  # expects YYYY-MM-DD
            except ValueError:
                print(f"Skipping row with bad date format: {row}")
                continue
            if dob_date.month == today.month and dob_date.day == today.day:
                matches.append((row["name"].strip(), row["email"].strip()))
    return matches


def already_sent_today(name):
    """Check sent_log.csv to avoid double-sending if the cron runs twice."""
    if not os.path.exists(LOG_PATH):
        return False
    today_str = date.today().isoformat()
    with open(LOG_PATH, newline="", encoding="utf-8") as f:
        for row in csv.reader(f):
            if len(row) >= 2 and row[0] == today_str and row[1] == name:
                return True
    return False


def log_sent(name, email):
    file_exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["date", "name", "email"])
        writer.writerow([date.today().isoformat(), name, email])


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
            img_data, maintype="image", subtype="png", cid="birthday_img"
        )
    else:
        print(f"WARNING: {IMAGE_PATH} not found - email will send without the image.")

    return msg


def send_email(msg):
    context = ssl.create_default_context()
    if SMTP_PORT == 465:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)


def main():
    if not SMTP_PASS:
        raise SystemExit("SMTP_PASS is not set. Set it as an environment variable before running.")

    birthdays = load_todays_birthdays(CSV_PATH)
    if not birthdays:
        print(f"No birthdays today ({date.today().isoformat()}).")
        return

    for name, email in birthdays:
        if already_sent_today(name):
            print(f"Already sent to {name} today, skipping.")
            continue
        if not email or email.strip().upper() == "TODO@IFOCUSSYSTEC.COM":
            print(f"Skipping {name} - no valid email set in employees.csv.")
            continue
        try:
            msg = build_message(name, email)
            send_email(msg)
            log_sent(name, email)
            print(f"Sent birthday email to {name} <{email}>")
        except Exception as e:
            print(f"FAILED to send to {name} <{email}>: {e}")


if __name__ == "__main__":
    main()
