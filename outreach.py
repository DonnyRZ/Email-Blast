#!/usr/bin/env python3
"""Send personalized outreach emails to hospitals listed in rs_online_1000.xlsx."""

import argparse
import mimetypes
import os
import re
import smtplib
import ssl
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from email.message import EmailMessage
from pathlib import Path
from string import Template
from typing import Dict, Iterable, List, Optional, Tuple

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
CELL_REF_PATTERN = re.compile(r"([A-Z]+)(\d+)")
DEFAULT_TEMPLATE_PATH = Path("templates") / "outreach_email.txt"
DEFAULT_RECIPIENTS = Path("test.xlsx")
TEST_RECIPIENTS = Path("test.xlsx")
DEFAULT_ENV_FILE = Path(".env")


def load_env_file(path: Path = DEFAULT_ENV_FILE) -> None:
    """Populate os.environ with key=value pairs from a .env file if present."""
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if not key or key in os.environ:
            continue
        value = value.strip().strip("'\"")
        os.environ[key] = value


def column_letters_to_index(letters: str) -> int:
    """Convert an Excel column label (e.g., 'AA') into a zero-based index."""
    index = 0
    for char in letters:
        index = index * 26 + (ord(char.upper()) - 64)
    return index - 1


def slugify(text: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")
    return slug


def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    try:
        raw = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    return ["".join(si.itertext()) for si in root.findall(f"{NS}si")]


def read_cell_value(cell: ET.Element, shared_strings: List[str]) -> str:
    cell_type = cell.get("t")
    if cell_type == "s":
        v = cell.find(f"{NS}v")
        if v is None or v.text is None:
            return ""
        idx = int(v.text)
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
    if cell_type == "inlineStr":
        inline = cell.find(f"{NS}is")
        if inline is None:
            return ""
        return "".join(inline.itertext())
    if cell_type == "b":
        v = cell.find(f"{NS}v")
        return "TRUE" if v is not None and v.text == "1" else "FALSE"
    v = cell.find(f"{NS}v")
    if v is not None and v.text is not None:
        return v.text
    return cell.findtext(f"{NS}t", default="") or ""


def iter_sheet_rows(workbook_path: Path) -> Iterable[Dict[str, str]]:
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook {workbook_path} does not exist.")
    with zipfile.ZipFile(workbook_path) as zf:
        shared = load_shared_strings(zf)
        try:
            sheet_xml = zf.read("xl/worksheets/sheet1.xml")
        except KeyError as exc:
            raise RuntimeError("Could not load the first worksheet from the Excel file.") from exc
        sheet_root = ET.fromstring(sheet_xml)
        sheet_data = sheet_root.find(f"{NS}sheetData")
        if sheet_data is None:
            return
        header_map = None
        for row in sheet_data.findall(f"{NS}row"):
            cells: Dict[int, str] = {}
            for cell in row.findall(f"{NS}c"):
                ref = cell.get("r")
                if not ref:
                    continue
                match = CELL_REF_PATTERN.match(ref)
                if not match:
                    continue
                col_index = column_letters_to_index(match.group(1))
                cells[col_index] = read_cell_value(cell, shared).strip()
            if not cells:
                continue
            if header_map is None:
                header_map = {idx: cells.get(idx, "") for idx in sorted(cells)}
                continue
            row_dict = {
                header_map[idx]: cells.get(idx, "").strip()
                for idx in header_map
                if header_map[idx]
            }
            if row_dict:
                yield row_dict


def build_context(row: Dict[str, str]) -> Dict[str, str]:
    context = {
        "hospital": (row.get("Rumah Sakit") or "").strip(),
        "province": (row.get("Provinsi") or "").strip(),
        "city": (row.get("Kab/Kota") or "").strip(),
        "address": (
            row.get("Alamat (List)")
            or row.get("Alamat (Profile)")
            or ""
        ).strip(),
        "phone": (
            row.get("Telepon (List)")
            or row.get("Telepon (Profile)")
            or ""
        ).strip(),
        "owner": (
            row.get("Pemilik (List)")
            or row.get("Kepemilikan")
            or ""
        ).strip(),
        "director": (row.get("Direktur") or "").strip(),
        "email": (row.get("Email Perusahaan") or "").strip(),
    }
    for key, value in row.items():
        slug = slugify(key)
        if slug and slug not in context:
            context[slug] = (value or "").strip()
    return context


def load_recipients(workbook_path: Path) -> List[Dict[str, str]]:
    recipients: List[Dict[str, str]] = []
    seen = set()
    for row in iter_sheet_rows(workbook_path):
        email = (row.get("Email Perusahaan") or "").strip()
        if not email or "@" not in email:
            continue
        key = email.lower()
        if key in seen:
            continue
        seen.add(key)
        recipients.append(build_context(row))
    return recipients


def load_attachments(directory: Optional[Path]) -> List[Tuple[str, bytes, str, str]]:
    """Load every regular file inside the directory into memory."""
    attachments: List[Tuple[str, bytes, str, str]] = []
    if not directory:
        return attachments
    directory = Path(directory)
    if not directory.exists() or not directory.is_dir():
        raise FileNotFoundError(f"Attachment directory {directory} does not exist or is not a folder.")
    for path in sorted(directory.iterdir()):
        if not path.is_file():
            continue
        mime_type, _ = mimetypes.guess_type(path.name)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        attachments.append((path.name, path.read_bytes(), maintype, subtype))
    return attachments


def format_sender_address(username: str, display_name: Optional[str]) -> str:
    if display_name:
        return f"{display_name} <{username}>"
    return username


def preview_messages(
    recipients: List[Dict[str, str]],
    subject_template: Template,
    body_template: Template,
    count: int,
) -> None:
    sample = recipients[:count]
    if not sample:
        print("No recipients with valid email addresses were found.")
        return
    for idx, recipient in enumerate(sample, start=1):
        subject = subject_template.safe_substitute(recipient)
        body = body_template.safe_substitute(recipient)
        print("=" * 70)
        print(f"[Preview {idx}] To: {recipient['email']}")
        print(f"Subject: {subject}")
        print(body)
        print("=" * 70)
    print(f"Previewed {len(sample)} message(s). Total sendable recipients: {len(recipients)}")


def send_messages(
    recipients: List[Dict[str, str]],
    subject_template: Template,
    body_template: Template,
    attachments: List[Tuple[str, bytes, str, str]],
    args: argparse.Namespace,
) -> None:
    connection_cls = smtplib.SMTP_SSL if args.implicit_tls else smtplib.SMTP
    ssl_context = ssl.create_default_context()
    total = len(recipients)
    successes = 0
    failures: List[Tuple[str, str]] = []

    def send_one(recipient: Dict[str, str]) -> None:
        msg = EmailMessage()
        msg["From"] = format_sender_address(args.smtp_user, args.from_name)
        msg["To"] = recipient["email"]
        msg["Subject"] = subject_template.safe_substitute(recipient)
        if args.reply_to:
            msg["Reply-To"] = args.reply_to
        msg.set_content(body_template.safe_substitute(recipient))
        for filename, data, maintype, subtype in attachments:
            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)

        retries = args.retries
        attempt = 0
        while True:
            attempt += 1
            try:
                with connection_cls(args.smtp_host, args.smtp_port, timeout=args.timeout) as client:
                    if not args.implicit_tls:
                        client.starttls(context=ssl_context)
                    client.login(args.smtp_user, args.smtp_password)
                    client.send_message(msg)
                return
            except Exception as exc:
                if attempt > retries:
                    raise
                backoff = args.retry_delay * attempt
                print(
                    f"Retrying {recipient['email']} after error: {exc}. Next attempt in {backoff:.1f}s.",
                    file=sys.stderr,
                )
                time.sleep(backoff)

    for idx, recipient in enumerate(recipients, start=1):
        try:
            send_one(recipient)
            successes += 1
            print(f"[{idx}/{total}] Sent {recipient['email']}")
        except Exception as exc:
            reason = str(exc)
            failures.append((recipient["email"], reason))
            print(f"[{idx}/{total}] Failed {recipient['email']}: {reason}", file=sys.stderr)
        if args.pause and idx < total:
            time.sleep(args.pause)

    print(f"Done. Success: {successes}, Failed: {len(failures)}.")
    if failures:
        print("Failures:")
        for email, reason in failures:
            print(f" - {email}: {reason}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send personalized outreach emails via mail.postale.io SMTP."
    )
    parser.add_argument("--subject", required=True, help="Subject template. Use $hospital, $city, etc.")
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE_PATH,
        help=f"Path to the email body template (default: {DEFAULT_TEMPLATE_PATH})",
    )
    parser.add_argument(
        "--recipients",
        type=Path,
        default=DEFAULT_RECIPIENTS,
        help=f"Path to the Excel workbook with contacts (default: {DEFAULT_RECIPIENTS})",
    )
    parser.add_argument("--smtp-user", default=os.getenv("SMTP_USER"), help="SMTP username (default: $SMTP_USER).")
    parser.add_argument(
        "--smtp-password",
        default=os.getenv("SMTP_PASSWORD"),
        help="SMTP password (default: $SMTP_PASSWORD).",
    )
    parser.add_argument("--smtp-host", default=os.getenv("SMTP_HOST", "mail.postale.io"))
    parser.add_argument("--smtp-port", type=int, default=int(os.getenv("SMTP_PORT", "587")))
    parser.add_argument(
        "--implicit-tls",
        action="store_true",
        help="Use implicit TLS on port 465 instead of STARTTLS on 587.",
    )
    parser.add_argument("--from-name", help="Optional display name for the From header.")
    parser.add_argument("--reply-to", help="Optional Reply-To address.")
    parser.add_argument("--pause", type=float, default=15.0, help="Seconds to pause between emails.")
    parser.add_argument("--retries", type=int, default=2, help="Number of times to retry a failed send per recipient.")
    parser.add_argument("--retry-delay", type=float, default=5.0, help="Base seconds to wait before retrying (multiplied per attempt).")
    parser.add_argument(
        "--use-test-data",
        action="store_true",
        help="Use test.xlsx instead of rs_online_1000.xlsx for safe testing.",
    )
    parser.add_argument(
        "--attachments-dir",
        type=Path,
        help="Attach every file in this directory to each email.",
    )
    parser.add_argument("--limit", type=int, help="Send to only the first N recipients.")
    parser.add_argument("--skip", type=int, default=0, help="Skip the first N recipients.")
    parser.add_argument("--dry-run", action="store_true", help="Preview emails without sending.")
    parser.add_argument(
        "--preview",
        type=int,
        default=3,
        help="Number of messages to show during a dry run.",
    )
    parser.add_argument("--timeout", type=int, default=30, help="SMTP connection timeout in seconds.")

    args = parser.parse_args()

    if not args.smtp_user or not args.smtp_password:
        parser.error("SMTP credentials are required. Supply --smtp-user/--smtp-password or set environment variables.")
    return args


def main() -> None:
    load_env_file()
    args = parse_args()

    if args.use_test_data:
        if not TEST_RECIPIENTS.exists():
            print(f"Test workbook {TEST_RECIPIENTS} does not exist.", file=sys.stderr)
            sys.exit(1)
        args.recipients = TEST_RECIPIENTS

    if not args.template.exists():
        print(f"Template file {args.template} does not exist.", file=sys.stderr)
        sys.exit(1)

    recipients = load_recipients(args.recipients)
    if args.skip:
        recipients = recipients[args.skip :]
    if args.limit is not None:
        recipients = recipients[: args.limit]

    if not recipients:
        print("No recipients with valid email addresses were found.", file=sys.stderr)
        sys.exit(1)

    try:
        attachments = load_attachments(args.attachments_dir)
    except FileNotFoundError as exc:
        print(exc, file=sys.stderr)
        sys.exit(1)

    subject_template = Template(args.subject)
    body_template = Template(args.template.read_text(encoding="utf-8-sig"))

    if args.dry_run:
        preview_messages(recipients, subject_template, body_template, args.preview)
        print(f"Dry run complete. Use --dry-run off to send {len(recipients)} messages.")
        return

    send_messages(recipients, subject_template, body_template, attachments, args)


if __name__ == "__main__":
    main()

