# iFocus Birthday Email Automation - Setup Guide

## 0. Install dependencies

```bash
pip install -r requirements.txt
```

This installs `openpyxl` (reads the employee `.xlsx` directly) and
`python-dotenv` (loads SMTP credentials from `.env`).

## 1. Set up employee data

- `employees.xlsx` in this repo is a **dummy example** (fake names/dates) that
  shows the required format - it's safe to commit and is tracked in git.
- Real employee data (names + birthdates are PII) must **never** be committed.
  Put it in **`employees.local.xlsx`** instead - that filename is gitignored,
  and the script automatically prefers it over `employees.xlsx` when it
  exists.
- Required columns (first row = header, any order): `name`, `email`, `dob`.
  `dob` can be a real Excel date cell or an ISO string (`YYYY-MM-DD`).
- Rows are skipped automatically (and logged, not silently dropped) if:
  - the email is blank or contains `TODO` (placeholder), or
  - the date can't be parsed.
- To point at a data file somewhere else entirely, set `EMPLOYEES_XLSX` in
  `.env` or the environment to an absolute path.

## 2. Get your SMTP details (from cPanel, not Roundcube)

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
5. Plug the host + port into `.env` (see next step).

If you don't have cPanel login access, whoever manages hosting/IT can pull
this in under 2 minutes from that same screen.

**You don't need any of this to test the email content** - see "Test without
SMTP credentials" below. Only fill in `SMTP_PASS` once you actually have
cPanel access.

## 3. Set credentials via `.env` (never hardcoded, never committed)

```bash
cp .env.example .env
```

Then edit `.env`:

```
SMTP_HOST=mail.ifocussystec.in
SMTP_PORT=465
SMTP_USER=hr.support@ifocussystec.com
SMTP_PASS=the-mailbox-password
```

`.env` is gitignored - it will never be committed. Leave `SMTP_PASS` blank
until you have real cPanel credentials; `--dry-run` works without it.

## 4. Add the birthday image

The cupcake graphic (`360_F_294637909_957UbRCZ8umRl6c6YzAcR78nAakfgSxf.jpg`)
is already in this repo and embedded inline in every email automatically.
If you ever need to swap it, update `IMAGE_PATH` in
`send_birthday_emails.py`.

## 5. Test without SMTP credentials (`--dry-run` and `--test-date`)

Before you have cPanel access (or any time you want to sanity-check the
template), use `--dry-run`. It builds the full email - including the
embedded cupcake image - and writes it to `dry_run_output/<date>_<name>.eml`
instead of sending anything over SMTP.

```bash
python3 send_birthday_emails.py --dry-run
```

Since nobody's birthday is likely to be today, combine it with `--test-date`
to simulate any date without touching your data file:

```bash
python3 send_birthday_emails.py --dry-run --test-date 2026-07-06
```

Open the resulting `.eml` file in Outlook/Thunderbird/Mail (or drag it into
Gmail) to see exactly what recipients would receive, image included.

Once you're happy with how it looks, run a **real dry run against real SMTP**
the same way, minus `--dry-run`, once `.env` has a real `SMTP_PASS`:

```bash
python3 send_birthday_emails.py --test-date 2026-07-06
```

This actually sends, so only do it once credentials are live and you're
ready to test end-to-end delivery.

## 6. Schedule it with cPanel Cron Jobs

*(Do this once you have server/cPanel access - not covered further here.)*

1. Upload this whole folder to the server, including a real `.env` (create
   it directly on the server - don't upload your local one over an insecure
   channel).
2. In cPanel, go to **Cron Jobs**.
3. Add a new cron job:
   - **Common Settings:** "Once per day" (or set manually to e.g. 9:00 AM ->
     Minute: `0`, Hour: `9`, everything else `*`)
   - **Command:**
     ```
     /usr/bin/python3 /home/yourcpaneluser/birthday_automation/send_birthday_emails.py
     ```
   - Adjust the path to wherever you actually upload this folder on the server.
4. Save. It'll now run automatically every day, silently, with no manual
   Roundcube compose needed.

## Notes

- **Logging:** every run writes to `birthday_automation.log` (gitignored) in
  addition to stdout, so after an unattended cron run you can check exactly
  what happened - who got emailed, who got skipped and why, any SMTP errors.
- **Duplicate prevention:** `sent_log.csv` (gitignored) is created
  automatically and prevents duplicate sends if the cron somehow runs twice
  in one day. It works the same way whether the underlying data source is
  `employees.xlsx` or `employees.local.xlsx`.
- **Retries:** SMTP sends automatically retry up to 2 times with exponential
  backoff (5s, then 10s) if the mail server is briefly unreachable, before
  being logged as a failure.
- BCC is currently set to `allemployees@ifocussystec.com`, matching what you're
  already doing manually - change `BCC_ADDRESS` in the script if that should differ.
- The image is embedded directly in the email (not remote-linked), so
  recipients won't see the "remote resources blocked" warning you saw in
  Roundcube.

## Security

- `.env`, `sent_log.csv`, `birthday_automation.log`, `dry_run_output/`, and
  `employees.local.xlsx` are all gitignored - only `.env.example` and the
  dummy `employees.xlsx` are meant to be committed.
- Never paste a real SMTP password into a chat, commit message, or any
  tracked file. It only ever belongs in the server's `.env`.
