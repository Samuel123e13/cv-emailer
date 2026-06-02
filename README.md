# CV Emailer

A small Windows desktop app to send your CV to many headhunters or companies at
once — but with a **separate, personalised email for each recipient**, so no one
sees anyone else's address and every message greets them by name.

![tabs: Recipients · Compose · Settings · Send](docs/screenshot.png)

## Features

- **Recipient list** — add/edit/remove contacts in a table, or **import from a
  CSV / Excel** file. Saved between sessions.
- **Personalised emails** — use `{name}`, `{first_name}`, `{company}` and
  `{email}` placeholders in your subject and body; each recipient gets their own.
- **CV attachment** — attach your CV (and any other files) once, sent to everyone.
- **Any email provider** — fully configurable SMTP, with one-click presets for
  Office 365 / university, Gmail, Outlook and Yahoo.
- **Safe testing** — "Send test to myself" before you send to real contacts.
- **Progress + log** — see each email succeed or fail, with a Stop button.
- **Password kept private** — stored in Windows Credential Manager, never in a
  file and never committed to git.

## Quick start

1. **Run it** — double-click `run.bat` (or `python app.py`).
2. **Settings tab** — pick a preset, enter your email + password, and click
   **Test connection**.
3. **Compose tab** — write your subject and body, and add your CV.
4. **Recipients tab** — add contacts or import a spreadsheet.
5. **Send tab** — click **Send test to myself**, check it looks right, then
   **Send to ALL recipients**.

### Recipient file format

A `.csv` or `.xlsx` with a header row. Columns are matched by name:

```csv
name,email,company
Jane Smith,jane.smith@example.com,Example Recruiters
```

See `sample_recipients.csv`.

## Finding recruiter emails (Hunter.io)

If you don't already have a list of contacts, the **Find emails (Hunter.io)**
button on the Recipients tab can look them up for you using
[Hunter.io](https://hunter.io)'s official API — a legitimate, Terms-of-Service
compliant service for finding *published* business emails (no scraping).

1. Sign up at [hunter.io](https://hunter.io) for a **free** API key
   (50 searches/month).
2. Paste it when the app asks; it's stored in Windows Credential Manager.
3. Click **Find emails (Hunter.io)**, enter company domains (one per line,
   e.g. `crowdstrike.com`), and click **Search**. Each company costs one search.
4. Rows that look like recruiters / HR are highlighted green and listed first.
   Select the ones you want and click **Add selected to recipients**.

> This finds business contacts companies have made public. Always send
> responsibly — relevant, individual applications, not indiscriminate bulk mail.

## Important: university / Microsoft 365 email

Many universities (including Microsoft 365 tenants) **disable "SMTP basic
authentication"**. If "Test connection" fails with a login error, that's almost
certainly why. Your options:

- Ask IT to enable **SMTP AUTH** for your mailbox, **or**
- Send from a personal **Gmail** account using an **App Password**
  (Google Account → Security → 2-Step Verification → App passwords), then use the
  Gmail preset in Settings.

## Sending responsibly

This tool is for **your own job search** — sending your CV to recruiters and
companies you intend to contact. Don't use it for unsolicited bulk mail. Most
providers also have daily sending limits (Gmail ~500/day), so keep lists
reasonable and leave the per-email delay enabled.

## Install the optional extras

```
pip install -r requirements.txt
```

- `keyring` — secure password storage (otherwise the password is session-only)
- `openpyxl` — import `.xlsx` files (CSV import works without it)

## Files

| File | Purpose |
|------|---------|
| `app.py` | The desktop UI |
| `sender.py` | Building and sending the emails over SMTP |
| `store.py` | Saving recipients, draft and settings; import; password storage |
| `run.bat` | Convenience launcher |

Your private data (`settings.json`, `recipients.json`, `draft.json`, your CV)
is git-ignored and stays on your machine.
