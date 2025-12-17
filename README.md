# Outreach Script

Python helper that turns the contact spreadsheets into personalized SMTP blasts from `donnyrp@steradian.co.id`. It uses Postale (`mail.postale.io`) and string templates, so everything stays lightweight and scriptable.

## Setup
1. Install Python 3.10+.
2. Copy `.env.example` to `.env` and fill in the real mailbox password. Use `SMTP_HOST=mail.postale.io` and `SMTP_PORT=587` unless Postale tells you otherwise.
3. Edit `templates/outreach_email.txt` if you want different copy (placeholders such as `$hospital`, `$city`, `$owner`, `$director`, etc.).

## Default test run
The script now points to `test.xlsx`, so you can fully exercise the pipeline without touching the real hospital list.

```powershell
# Preview the message without sending anything
python outreach.py --subject 'Test send to tal' --from-name 'Donny at San' --dry-run

# Send the actual test email (uses creds from .env)
python outreach.py --subject 'Test send to tal' --from-name 'Donny at San' --reply-to 'donnyrp@steradian.co.id' --preview 0
```

## Going live with rs_online_1000.xlsx
1. Always dry-run first so you can double-check personalization.
2. Remove `--dry-run` only when you’re ready to deliver.

```powershell
# Dry-run against the production list
python outreach.py --subject 'Steradian support for $hospital' --from-name 'Donny at Steradian' --recipients rs_online_1000.xlsx --dry-run --preview 3

# Send to the first 100 contacts with a short pause
python outreach.py --subject 'Steradian support for $hospital' --from-name 'Donny at Steradian' --recipients rs_online_1000.xlsx --reply-to 'donnyrp@steradian.co.id' --limit 100 --pause 2.0
```

## Handy switches
| Flag | Meaning |
| --- | --- |
| `--subject` | Subject template. Use placeholders like `$hospital`, `$city`, `$province`. |
| `--template PATH` | Alternate body template file (defaults to `templates/outreach_email.txt`). |
| `--recipients FILE` | Excel workbook. Defaults to `test.xlsx`; point it at `rs_online_1000.xlsx` when going live. |
| `--pause` | Seconds to wait between messages. Helpful for rate limits. |
| `--limit / --skip` | Batch sends by slicing the recipient list. |
| `--dry-run` | Preview without touching SMTP. Combine with `--preview N` to control how many samples you see. |
| `--implicit-tls` | Use port 465 instead of STARTTLS on 587 if Postale ever asks for it. |

## Tips & gotchas
- Leave `.env` out of git (already handled by `.gitignore`). Never commit real credentials.
- Missing placeholders resolve to empty strings, so optional data won’t break the send.
- If you see `CERTIFICATE_VERIFY_FAILED`, confirm `SMTP_HOST` is `mail.postale.io`.
- For any spreadsheet row with a blank `Email Perusahaan`, the script simply skips it.
