# iFocus Birthday Email Automation + HR Portal

Two pieces that share one employee database (`employees.db`, SQLite):

- **`send_birthday_emails.py`** - run daily (cron/Task Scheduler) to email
  anyone whose birthday is today.
- **HR Portal** (`portal_app.py` / `run_portal.py`) - a small web page HR
  opens from her browser (over the office WiFi/LAN) to view, add, edit,
  delete employees, or bulk-upload an Excel file. Edits she makes are picked
  up automatically by the next scheduled email run - no code changes, no
  manually handing you an Excel file.

## 0. Install dependencies

```bash
pip install -r requirements.txt
```

Installs `openpyxl` (Excel import/export), `python-dotenv` (loads `.env`),
`Flask` (the portal), and `waitress` (production web server for the portal).

## 1. Employee data lives in `employees.db`

The first time either script runs, it creates `employees.db` next to it and,
if that DB is brand new, automatically imports whatever it finds in
`employees.local.xlsx` (falling back to the dummy `employees.xlsx`) so no
existing data is lost in the switch to a database.

From then on, **`employees.db` is the source of truth** - edit employee data
through the HR Portal (below), not by hand-editing Excel files. Excel is now
just an import/export format:

- **Upload** an `.xlsx` in the portal to bulk-add or bulk-update employees.
  Required columns (header row, any order): `name`, `email`, `dob`
  (`YYYY-MM-DD` or a real Excel date cell). Matching is by email - existing
  employees are updated, new emails are added, nothing is ever deleted by an
  upload.
- **Export** downloads the current database as an `.xlsx`, e.g. for a backup
  or to hand off to payroll.

`employees.db` contains real PII and is gitignored - never commit it.

## 2. Configure SMTP and the portal password

```bash
cp .env.example .env
```

Edit `.env`:

```
SMTP_HOST=mail.ifocussystec.in
SMTP_PORT=465
SMTP_SECURE=true
SMTP_USER=itsupport@ifocussystec.com
SMTP_PASSWORD=the-real-mailbox-password
SMTP_FROM_NAME=HR Support

PORTAL_PASSWORD=pick-a-shared-password-for-hr
SECRET_KEY=generate-with-the-command-below
```

Generate a `SECRET_KEY` once and keep it stable (changing it logs everyone
out of the portal):

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

`.env` is gitignored - it will never be committed. `--dry-run` on the email
script works without `SMTP_PASSWORD` set at all.

**`PORTAL_PASSWORD` is the only thing standing between "on the office WiFi"
and "can view/edit every employee's name, email and birthdate."** Set it to
something real before telling HR the portal's address, and only share it
with her directly (not over the same channel as the URL).

## 3. Run the HR Portal

**Local testing** (just on your own machine):

```bash
python portal_app.py
```

Open http://localhost:5000, log in with `PORTAL_PASSWORD`.

**Always-on hosting on the LAN** (so HR can reach it any time, not just when
your laptop happens to be open): run `run_portal.py` on whatever machine is
meant to stay on - server, NAS, an old desktop left running, etc. It uses
`waitress` (a proper production server, unlike Flask's built-in dev server)
and binds to all network interfaces:

```bash
python run_portal.py
```

It prints the LAN URL to use, e.g. `http://192.168.1.42:5000`. Anyone on the
same WiFi/LAN can open that URL from a laptop or phone browser and log in
with `PORTAL_PASSWORD`.

To keep it running permanently in the background:

- **Windows** - use Task Scheduler: create a task that runs
  `pythonw.exe run_portal.py` at system startup, or install it as a proper
  Windows service with [NSSM](https://nssm.cc/) (`nssm install HRPortal
  "C:\path\to\python.exe" "C:\path\to\run_portal.py"`).
- **Linux/NAS with systemd** - create `/etc/systemd/system/hr-portal.service`:
  ```ini
  [Unit]
  Description=iFocus HR Portal
  After=network.target

  [Service]
  WorkingDirectory=/path/to/email-automater
  ExecStart=/usr/bin/python3 run_portal.py
  Restart=on-failure

  [Install]
  WantedBy=multi-user.target
  ```
  Then `sudo systemctl enable --now hr-portal`.

**Firewall:** whichever machine hosts the portal needs to allow inbound
connections on port 5000 from the local network (Windows Defender Firewall
will prompt the first time you run it - allow it for "Private" networks
only, not "Public").

## 4. Run the birthday email sender

```bash
python send_birthday_emails.py
python send_birthday_emails.py --dry-run
python send_birthday_emails.py --dry-run --test-date 2026-07-06
```

`--dry-run` builds the full email - including the embedded cupcake image -
and writes it to `dry_run_output/<date>_<name>.eml` instead of sending
anything over SMTP. Open the `.eml` in Outlook/Thunderbird/Mail (or drag it
into Gmail) to preview exactly what a recipient would see.

Since nobody's birthday is likely to be today, combine with `--test-date` to
simulate any date without touching real data.

### Schedule it daily

**Windows (Task Scheduler):**
Create a daily trigger (e.g. 9:00 AM) that runs:
```
C:\path\to\python.exe C:\path\to\send_birthday_emails.py
```

**Linux/NAS (cron):**
```bash
# crontab -e
0 9 * * * /usr/bin/python3 /path/to/send_birthday_emails.py >> /path/to/cron.log 2>&1
```

If the portal is hosted on the same always-on machine, it makes sense to
schedule the email sender there too, so it always sees HR's latest edits.

## 5. Add the birthday image

The cupcake graphic (`360_F_294637909_957UbRCZ8umRl6c6YzAcR78nAakfgSxf.jpg`)
is already in this repo and embedded inline in every email automatically.
To swap it, update `IMAGE_PATH` in `send_birthday_emails.py`.

## Notes

- **Logging:** every email run writes to `birthday_automation.log`
  (gitignored) in addition to stdout - who got emailed, who got skipped and
  why, any SMTP errors.
- **Duplicate prevention:** `sent_log.csv` (gitignored) prevents duplicate
  sends if the scheduler somehow runs twice in one day.
- **Retries:** SMTP sends automatically retry up to 2 times with exponential
  backoff (5s, then 10s) before being logged as a failure.
- BCC is currently set to `ramkirangaruda2006@gmail.com` for testing -
  change `BCC_ADDRESS` in `send_birthday_emails.py` back to
  `allemployees@ifocussystec.com` before going live.
- The birthday image is embedded directly in the email (not remote-linked),
  so recipients won't see a "remote resources blocked" warning.
- Rows are skipped automatically (and reported, never silently dropped) on
  Excel import if the email is blank, contains `TODO`, isn't a valid email
  format, or the date can't be parsed.

## Security

- `.env`, `employees.db`, `sent_log.csv`, `birthday_automation.log`, and
  `dry_run_output/` are all gitignored - only `.env.example` and the dummy
  `employees.xlsx` are meant to be committed.
- Never paste the real SMTP password or `PORTAL_PASSWORD` into a chat,
  commit message, or any tracked file.
- The portal has no per-user accounts - it's one shared password gating
  everyone who can reach it on the LAN. Treat `PORTAL_PASSWORD` like a
  mailbox password: share it verbally/directly with HR, not by email or in
  a shared doc next to the portal's URL.
- The portal is designed for **LAN-only** access (office WiFi). Don't port-
  forward it to the public internet without adding real authentication and
  HTTPS first - a single shared password over plain HTTP is fine on a
  trusted internal network, not on the open internet.
