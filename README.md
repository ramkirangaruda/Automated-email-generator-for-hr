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

## 2. Enable Gmail "Less secure apps" access

The script uses Gmail's SMTP server (`smtp.gmail.com:587`). By default, Gmail
blocks less-secure apps, so you need to enable this once:

1. Open https://myaccount.google.com/u/0/security
2. On the left, click **Security**
3. Scroll down to **"Less secure app access"** and toggle it **ON**
4. (If you don't see this option, 2-Step Verification may not be enabled; Gmail
   recommends using App Passwords instead — but for this use case, "Less secure
   apps" is simpler)

Once enabled, the mailbox password works as-is for SMTP auth.

**You don't need to do this yet** — `--dry-run` works without Gmail credentials.
Only enable this once you're ready to test real delivery.

## 3. Set credentials via `.env` (never hardcoded, never committed)

```bash
cp .env.example .env
```

Then edit `.env` and fill in `SMTP_PASS` with the mailbox password:

```
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=hr.support@ifocussystec.com
SMTP_PASS=the-gmail-mailbox-password
```

`.env` is gitignored - it will never be committed. Leave `SMTP_PASS` blank
until you've enabled "Less secure apps"; `--dry-run` works without it.

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

## 6. Deploy and schedule (optional)

Once you're happy with local testing:

1. Upload this repo to your server (any Linux/Unix machine with Python 3.8+).
2. On the server, create `.env` with the Gmail credentials (never upload your
   local `.env` — create it fresh on the server for security).
3. Install dependencies: `pip install -r requirements.txt`
4. Test it once manually: `python3 send_birthday_emails.py`
5. Set up a cron job to run it daily:
   ```bash
   # crontab -e
   # Add a line like:
   0 9 * * * /path/to/repo/send_birthday_emails.py >> /path/to/repo/cron.log 2>&1
   ```
   This runs at 9:00 AM every day. Adjust the time as needed.

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
- Never paste a real Gmail password into a chat, commit message, or any
  tracked file. The password only belongs in the server's `.env` (or locally
  if you're testing).
- If you enable "Less secure apps" on the Gmail account, make sure only
  trusted machines have the `.env` file with credentials.
