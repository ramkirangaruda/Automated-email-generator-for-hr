# iFocus Birthday Email Automation - Setup Guide

## 1. Get your SMTP details (from cPanel, not Roundcube)

Roundcube itself won't show you the raw SMTP host/port - that's set by the
mail server admin. Since your webmail URL was `webmail.ifocussystec.in/cpsess.../roundcube`,
your mailbox is cPanel-hosted, so:

1. Log into **cPanel** directly (usually `ifocussystec.in/cpanel` or ask whoever
   set up hosting for the login link).
2. Go to **Email Accounts**.
3. Find the sending account (e.g. `hr.support@ifocussystec.com`) and click
   **Connect Devices** (sometimes labeled "Set Up Mail Client").
4. It will show **Manual Settings** with something like:
   - Incoming/Outgoing Server: `mail.ifocussystec.in` (or `ifocussystec.in`)
   - SMTP Port: `465` (SSL) or `587` (STARTTLS)
   - Username: the full email address
   - Password: the mailbox password
5. Plug the host + port into the script's `SMTP_HOST` / `SMTP_PORT`.

If you don't have cPanel login access, whoever manages hosting/IT can pull
this in under 2 minutes from that same screen.

## 2. Set credentials as environment variables (don't hardcode the password)

```bash
export SMTP_HOST="mail.ifocussystec.in"
export SMTP_PORT="465"
export SMTP_USER="hr.support@ifocussystec.com"
export SMTP_PASS="the-mailbox-password"
```

## 3. Add the real employee data

- `employees.csv` currently has the 14 names you gave me, with dates converted
  to `YYYY-MM-DD` format and `email` set to `TODO@ifocussystec.com` as a placeholder.
- When you build the real Excel sheet, just make sure it has 3 columns:
  `name`, `email`, `dob` (dob as `YYYY-MM-DD`) - then **Save As > CSV** and
  replace `employees.csv`. Any spreadsheet tool can export CSV directly.
- Rows with a `TODO` email are automatically skipped (won't error, just logged).

## 4. Add the birthday image

Save the actual cupcake graphic (the one that appears in every HR birthday
email) as `birthday_image.png` in this same folder. Easiest way to grab it:
open one of the birthday emails in Roundcube -> right-click the image ->
"Save image as" -> save it here as `birthday_image.png`.

## 5. Test it manually first

```bash
cd birthday_automation
python3 send_birthday_emails.py
```

Since nobody's birthday is likely today, temporarily edit one row's `dob` in
`employees.csv` to today's date, run the script, confirm the email arrives
correctly, then change it back.

## 6. Schedule it with cPanel Cron Jobs

1. In cPanel, go to **Cron Jobs**.
2. Add a new cron job:
   - **Common Settings:** "Once per day" (or set manually to e.g. 9:00 AM ->
     Minute: `0`, Hour: `9`, everything else `*`)
   - **Command:**
     ```
     SMTP_PASS="the-password" SMTP_USER="hr.support@ifocussystec.com" SMTP_HOST="mail.ifocussystec.in" SMTP_PORT="465" /usr/bin/python3 /home/yourcpaneluser/birthday_automation/send_birthday_emails.py >> /home/yourcpaneluser/birthday_automation/cron.log 2>&1
     ```
   - Adjust the path to wherever you actually upload this folder on the server.
3. Save. It'll now run automatically every day, silently, with no manual
   Roundcube compose needed.

## Notes

- `sent_log.csv` is created automatically and prevents duplicate sends if the
  cron somehow runs twice in one day.
- BCC is currently set to `allemployees@ifocussystec.com`, matching what you're
  already doing manually - change `BCC_ADDRESS` in the script if that should differ.
- The image is embedded directly in the email (not remote-linked), so
  recipients won't see the "remote resources blocked" warning you saw in
  Roundcube.
