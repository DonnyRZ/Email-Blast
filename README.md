# Mass Outreach Helper

This repository now includes a self-contained Python tool for sending personalized outreach emails from `donnyrp@steradian.co.id` (Postale.io SMTP) to the hospital list in `rs_online_1000.xlsx`.

## Features
- Pure-Python Excel reader (no third-party packages) tuned for `rs_online_1000.xlsx`.
- Recipient normalization with automatic deduplication by `Email Perusahaan`.
- Text template rendering via `string.Template` with contextual placeholders (`$hospital`, `$city`, `$province`, `$owner`, etc.).
- CLI controls for dry-runs, batching (`--pause`), limiting/offsetting recipients, and TLS mode selection.
- Works with environment variables (`SMTP_USER`, `SMTP_PASSWORD`, `SMTP_HOST`, `SMTP_PORT`) so credentials never live in the repository.

## Prerequisites
- Python 3.10+ available on your machine (the standard Windows Store build works).
- The Postale SMTP password for `donnyrp@steradian.co.id`.

## Quick Start
```powershell
# (optional) load credentials into the environment for convenience
$env:SMTP_USER = "donnyrp@steradian.co.id"
$env:SMTP_PASSWORD = "<your-mailbox-password>"

# Dry-run to preview the first few personalized messages
python outreach.py --subject 'Collaboration with $hospital' --from-name "Donny at Steradian" --dry-run

# Actually send to the first 50 test entries with 2 seconds between messages
python outreach.py `
  --subject 'Steradian support for $hospital' `
  --from-name "Donny at Steradian" `
  --reply-to "donnyrp@steradian.co.id" `
  --limit 50 `
  --pause 2.0
```

Prefer dot-env files? Copy `.env.example` to `.env`, fill in the real credentials, and the script will pick them up automatically before parsing CLI flags.

### Testing with mock data
To experiment without touching the production list, run the script against `test.xlsx`:

```powershell
python outreach.py --subject 'Test send to $hospital' --dry-run
```

`test.xlsx` mirrors the real schema but points to dummy recipients so you can validate personalization, throttling, and SMTP connectivity safely. When you are ready to send to the full hospital list, simply add `--recipients rs_online_1000.xlsx`. If you rename the mock file, you can still use `--recipients` to point at it explicitly or fall back to `--use-test-data`.

### Command reference
| Flag | Description |
| --- | --- |
| `--subject` | Required subject template. Placeholders such as `$hospital`, `$city`, `$province`, `$owner`, `$director`, `$email`, plus slugified forms of every Excel column (e.g. `$rumah_sakit`, `$kab_kota`, `$email_perusahaan`). |
| `--template` | Email body path (defaults to `templates/outreach_email.txt`). Use any placeholders from above. |
| `--recipients` | Excel file path (`test.xlsx` by default while we are in testing mode). |
| `--smtp-user / --smtp-password` | Override credentials if not set via env vars. Host/port default to `mail.postale.io:587`. |
| `--implicit-tls` | Switches to implicit TLS (port 465). Otherwise STARTTLS on 587 is used. |
| `--from-name`, `--reply-to` | Customize headers. |
| `--limit`, `--skip` | Control slices of the recipient list. Useful for batching across multiple days. |
| `--pause` | Seconds to sleep between sends (prevents hitting rate limits). |
| `--use-test-data` | Swap to `test.xlsx` for safe dry runs without touching the production recipient list. |
| `--dry-run`, `--preview` | Preview without connecting to SMTP. |

## Template customization
Edit `templates/outreach_email.txt` or create additional files that follow the same placeholder syntax. Everything between `$` braces maps to either:
- Canonical keys we add (`hospital`, `city`, `province`, `address`, `phone`, `owner`, `director`, `email`).
- Slugified column names from the spreadsheet (`Alamat (Profile)` -> `$alamat_profile`, etc.).

Because we use `Template.safe_substitute`, missing placeholders simply become empty strings, which is helpful for optional data points.

## Sending tips
- Use `--dry-run` before every campaign to double-check personalization.
- Break large blasts into batches with `--limit`/`--skip` and keep at least 1-2 seconds between sends.
- Monitor Postale's outbound limits; adjust `--pause` accordingly.
- Keep the Excel file updated; the script automatically deduplicates repeated email addresses, so you can rerun without worrying about duplicates.

## Troubleshooting
- `Workbook ... does not exist`: verify `rs_online_1000.xlsx` lives next to `outreach.py` or pass `--recipients`.
- `SMTPAuthenticationError`: ensure the correct mailbox password (normal password, not an app password) and confirm no 2FA restrictions are blocking SMTP.
- Empty preview: the spreadsheet rows you selected might not have `Email Perusahaan` filled. Add addresses or adjust the list.

