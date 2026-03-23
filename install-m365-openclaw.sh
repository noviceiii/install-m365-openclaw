#!/usr/bin/env bash
# =============================================================================
# OpenClaw M365 Graph Skill - Installer 
# =============================================================================
# Features:
# - Asks for OpenClaw base directory and M365 user email
# - Creates proper package structure: src/m365_openclaw/__init__.py
# - Installs as editable package with correct setup.py
# - Full M365Client: Mail, Calendar, Contacts, OneDrive, OneNote, Excel, Word, PPT
# - All 13 CLI commands implemented in __main__.py
# - SKILL.md generated for automatic OpenClaw agent discovery
# - M365_USER_EMAIL stored in .env for correct daemon-mode API routing
# - main_resource set so /me/ resolves under client-credentials flow
# - Verifies package import after installation
# - PATH added permanently (only once)
# - Uses venv Python and pip explicitly
# - set -euo pipefail safe (PYTHONPATH with :- fallback)
# =============================================================================

set -euo pipefail

# Default values
DEFAULT_OPENCLAW_DIR="${HOME}/.openclaw"
DEFAULT_M365_USER="admin@yourcompany.onmicrosoft.com"

echo "=== OpenClaw M365 Graph Skill Installer - Reliable Version ==="

# Ask for OpenClaw base directory
read -p "OpenClaw base directory [default: ${DEFAULT_OPENCLAW_DIR}]: " OPENCLAW_DIR
OPENCLAW_DIR="${OPENCLAW_DIR:-${DEFAULT_OPENCLAW_DIR}}"

# Ask for M365 user email (guidance only)
read -p "M365 account email for authentication [default: ${DEFAULT_M365_USER}]: " M365_USER
M365_USER="${M365_USER:-${DEFAULT_M365_USER}}"

SKILL_NAME="m365-graph"
INSTALL_DIR="${OPENCLAW_DIR}/skills/${SKILL_NAME}"
VENV_DIR="${INSTALL_DIR}/venv"
CLI_NAME="m365"
CONFIG_FILE="${INSTALL_DIR}/.env"
TOKEN_CACHE="${OPENCLAW_DIR}/credentials/m365_token_cache.bin"
SRC_DIR="${INSTALL_DIR}/src"

echo ""
echo "Installation summary:"
echo "  OpenClaw base:     ${OPENCLAW_DIR}"
echo "  Install path:      ${INSTALL_DIR}"
echo "  Venv path:         ${VENV_DIR}"
echo "  Suggested account: ${M365_USER}"
echo ""

# 1. System dependencies (sudo required)
sudo apt update -qq
sudo apt install -y python3 python3-venv python3-pip curl git

# 2. Create directories
mkdir -p "${INSTALL_DIR}"/{bin,src/m365_openclaw}
mkdir -p "${OPENCLAW_DIR}/credentials"
mkdir -p "${HOME}/.local/bin"

# 3. Create virtual environment and install packages
python3 -m venv "${VENV_DIR}"
"${VENV_DIR}/bin/pip" install --upgrade pip
"${VENV_DIR}/bin/pip" install \
    O365 msal msal_extensions python-dotenv \
    python-docx python-pptx openpyxl

# 4. Create proper package structure and setup.py
touch "${SRC_DIR}/m365_openclaw/__init__.py"

cat > "${INSTALL_DIR}/setup.py" << 'EOF'
from setuptools import setup, find_packages

setup(
    name='m365_openclaw',
    version='0.1.0',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
)
EOF

# Install editable package
cd "${INSTALL_DIR}"
echo "Installing m365_openclaw package in editable mode..."
"${VENV_DIR}/bin/pip" install -e .

# Verify import works
echo "Verifying package import..."
if "${VENV_DIR}/bin/python" -c "import m365_openclaw; print('SUCCESS: m365_openclaw imported')" 2>/dev/null; then
    echo "Package verification successful."
else
    echo "ERROR: Package import failed. Check setup.py and src/m365_openclaw/__init__.py"
    exit 1
fi

# 5. Create CLI wrapper
cat > "${INSTALL_DIR}/bin/m365" << EOF
#!/usr/bin/env bash
export PYTHONPATH="${SRC_DIR}:\${PYTHONPATH}"
exec "${VENV_DIR}/bin/python" -m m365_openclaw "\$@"
EOF

chmod +x "${INSTALL_DIR}/bin/m365"

ln -sf "${INSTALL_DIR}/bin/m365" "${HOME}/.local/bin/${CLI_NAME}"

# Add PATH permanently (only once)
if ! grep -q "${HOME}/.local/bin" "${HOME}/.bashrc"; then
    echo 'export PATH="$HOME/.local/bin:$PATH"' >> "${HOME}/.bashrc"
    echo "Added PATH entry to ~/.bashrc"
fi
export PATH="${HOME}/.local/bin:$PATH"

# 6. Create config file
cat > "${CONFIG_FILE}" << EOF
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-app-client-id-here
CLIENT_SECRET=your-app-client-secret-here
TOKEN_CACHE_PATH=${TOKEN_CACHE}
M365_USER_EMAIL=${M365_USER}
EOF

# 7. Entra ID instructions
echo ""
echo "=== Entra ID App Registration (one-time) ==="
echo "1. Go to https://entra.microsoft.com → App registrations → New registration"
echo "   Name: OpenClaw-M365-Agent"
echo "   Supported account types: Accounts in this organizational directory only"
echo "   Redirect URI: leave blank"
echo "2. API permissions → Microsoft Graph"
echo "   Application: Mail.ReadWrite.All, Mail.Send, Calendars.ReadWrite.All,"
echo "                Files.ReadWrite.All, Contacts.ReadWrite"
echo "   Delegated:   Notes.ReadWrite.All, offline_access"
echo "   → Grant admin consent"
echo "3. Certificates & secrets → New client secret → copy Value"
echo ""
read -p "Press ENTER when you have TENANT_ID, CLIENT_ID, CLIENT_SECRET..."

read -p "TENANT_ID: " TENANT_ID
read -p "CLIENT_ID: " CLIENT_ID
read -s -p "CLIENT_SECRET: " CLIENT_SECRET
echo ""

sed -i "s|your-tenant-id-here|${TENANT_ID}|" "${CONFIG_FILE}"
sed -i "s|your-app-client-id-here|${CLIENT_ID}|" "${CONFIG_FILE}"
sed -i "s|your-app-client-secret-here|${CLIENT_SECRET}|" "${CONFIG_FILE}"

# 8. Create modules inside package
cat > "${SRC_DIR}/m365_openclaw/core.py" << 'PYEOF'
"""
core.py – Microsoft 365 client for OpenClaw agents.

Supports: Mail, Calendar, Contacts, OneDrive, OneNote, Excel, Word, PowerPoint.
Uses client-credentials (daemon/application) flow via the O365 library + MSAL.

Important: When using application permissions (client_credentials flow), all
Graph API calls must target a specific user.  Set M365_USER_EMAIL in .env so
that main_resource resolves /me/ to /users/<email>/ automatically.
"""

import json
import os
import re
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

from dotenv import load_dotenv
from O365 import Account
from O365.utils import FileSystemTokenBackend

load_dotenv()


def _parse_excel_range(range_str):
    """Return (rows, cols) from a range such as 'A1' or 'A1:C3'."""
    m = re.match(r'^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$', range_str.strip())
    if not m:
        return 1, 1
    sc, sr, ec, er = m.group(1), m.group(2), m.group(3), m.group(4)
    if ec is None:
        return 1, 1

    def col_num(col):
        n = 0
        for ch in col.upper():
            n = n * 26 + (ord(ch) - ord('A') + 1)
        return n

    return int(er) - int(sr) + 1, col_num(ec) - col_num(sc) + 1


class M365Client:
    def __init__(self):
        self.tenant_id = os.getenv("TENANT_ID")
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.token_cache_path = os.getenv("TOKEN_CACHE_PATH")
        # main_resource ensures /me/ resolves to the licensed user under
        # client-credentials (application permissions) flow.
        self.user = os.getenv("M365_USER_EMAIL") or "me"

        if not all([self.tenant_id, self.client_id, self.client_secret]):
            raise EnvironmentError(
                "Missing required credentials. "
                "Edit ~/.openclaw/skills/m365-graph/.env and set "
                "TENANT_ID, CLIENT_ID, and CLIENT_SECRET."
            )

        credentials = (self.client_id, self.client_secret)
        token_backend = FileSystemTokenBackend(
            token_path=Path(self.token_cache_path)
        )

        self.account = Account(
            credentials=credentials,
            auth_flow_type='credentials',
            tenant_id=self.tenant_id,
            token_backend=token_backend,
            main_resource=self.user,
        )

        if not self.account.is_authenticated:
            print("Authenticating with Microsoft 365...", file=sys.stderr)
            self.account.authenticate(
                scopes=['https://graph.microsoft.com/.default']
            )

    # ── internal helpers ──────────────────────────────────────────────────────

    def _get_drive(self):
        storage = self.account.storage()
        return storage.get_default_drive(request_if_none=True)

    def _access_token(self):
        """Return a valid access token string from the cached token."""
        token = self.account.connection.token_backend.token
        return token.get('access_token', '')

    def _graph_patch(self, url, data):
        """HTTP PATCH to the Microsoft Graph API with JSON body."""
        import requests as req
        headers = {
            'Authorization': f'Bearer {self._access_token()}',
            'Content-Type': 'application/json',
        }
        resp = req.patch(url, headers=headers, json=data, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    # ── Mail ──────────────────────────────────────────────────────────────────

    def send_mail(self, to_address, subject, body):
        """Send an email to any recipient."""
        m = self.account.mailbox().new_message()
        m.to.add(to_address)
        m.subject = subject
        m.body = body
        m.send()
        return f"Email sent to {to_address}"

    def list_mail(self, limit=20):
        """Return recent messages from the inbox."""
        messages = self.account.mailbox().inbox_folder().get_messages(limit=limit)
        result = []
        for msg in messages:
            result.append({
                'subject': msg.subject or '(no subject)',
                'from': str(msg.sender),
                'date': msg.received.isoformat() if msg.received else None,
                'is_read': msg.is_read,
            })
        return result

    # ── Calendar ──────────────────────────────────────────────────────────────

    def get_calendar_events(self, days=7):
        """Return upcoming calendar events within the next *days* days."""
        schedule = self.account.schedule()
        calendar = schedule.get_default_calendar()
        events = calendar.get_events(include_recurring=False, limit=100)
        now_utc = datetime.now(timezone.utc)
        cutoff = now_utc + timedelta(days=days)
        upcoming = []
        for event in events:
            start = event.start
            if start and start.tzinfo is None:
                start = start.replace(tzinfo=timezone.utc)
            if start and now_utc <= start <= cutoff:
                upcoming.append({
                    'subject': event.subject or '(no subject)',
                    'start': start.isoformat(),
                    'end': event.end.isoformat() if event.end else None,
                    'location': str(event.location) if event.location else None,
                })
        return upcoming

    def create_calendar_event(self, subject, start_iso, end_iso, body="", location=""):
        """Create a calendar event. start/end must be ISO 8601 strings."""
        schedule = self.account.schedule()
        calendar = schedule.get_default_calendar()
        event = calendar.new_event()
        event.subject = subject
        event.body = body
        if location:
            event.location = location
        event.start = datetime.fromisoformat(start_iso)
        event.end = datetime.fromisoformat(end_iso)
        event.save()
        return f"Event '{subject}' created"

    # ── Contacts ──────────────────────────────────────────────────────────────

    def list_contacts(self, limit=100):
        """List contacts from the personal address book."""
        address_book = self.account.address_book()
        contacts = address_book.get_contacts(limit=limit)
        result = []
        for contact in contacts:
            emails = []
            if hasattr(contact, 'emails') and contact.emails:
                emails = [str(e) for e in contact.emails]
            phones = []
            if hasattr(contact, 'business_phones') and contact.business_phones:
                phones = list(contact.business_phones)
            result.append({
                'name': contact.full_name or '(no name)',
                'emails': emails,
                'phones': phones,
            })
        return result

    def create_contact(self, given_name, surname, email=None, phone=None):
        """Create a new contact in the personal address book."""
        address_book = self.account.address_book()
        contact = address_book.new_contact()
        contact.given_name = given_name
        contact.surname = surname
        if email:
            contact.emails.add(email)
        if phone:
            contact.business_phones.append(phone)
        contact.save()
        return f"Contact '{given_name} {surname}' created"

    # ── OneDrive ─────────────────────────────────────────────────────────────

    def onedrive_list(self, folder_path="/"):
        """List files and folders in OneDrive."""
        drive = self._get_drive()
        if folder_path in ("/", ""):
            folder = drive.get_root_folder()
        else:
            folder = drive.get_item_by_path(folder_path)
        result = []
        for item in folder.get_items():
            result.append({
                'name': item.name,
                'type': 'folder' if item.is_folder else 'file',
                'size': item.size if not item.is_folder else None,
                'modified': item.modified.isoformat() if item.modified else None,
            })
        return result

    def onedrive_upload(self, local_path, remote_path):
        """Upload a local file to OneDrive, overwriting if it already exists."""
        drive = self._get_drive()
        remote = Path(remote_path)
        parent_str = str(remote.parent)
        if parent_str in ("/", "."):
            folder = drive.get_root_folder()
        else:
            folder = drive.get_item_by_path(parent_str)
        uploaded = folder.upload_file(item=local_path, item_name=remote.name)
        return f"Uploaded to {remote_path}" if uploaded else "Upload failed"

    def onedrive_download(self, remote_path, local_path):
        """Download a file from OneDrive to a local path."""
        drive = self._get_drive()
        item = drive.get_item_by_path(remote_path)
        local = Path(local_path)
        item.download(to_path=str(local.parent), name=local.name)
        return f"Downloaded to {local_path}"

    # ── OneNote ───────────────────────────────────────────────────────────────

    def onenote_create_page(self, notebook_name, section_name, title, html_content):
        """
        Create a OneNote page in the given notebook and section.
        Note: Requires the delegated permission Notes.ReadWrite.All.
        Application permissions do not support OneNote as of early 2025.
        """
        try:
            onenote = self.account.onenote()
            notebooks = list(onenote.list_notebooks())
            notebook = next(
                (nb for nb in notebooks if nb.name.lower() == notebook_name.lower()),
                None,
            )
            if not notebook:
                notebook = onenote.create_notebook(name=notebook_name)
            sections = list(notebook.list_sections())
            section = next(
                (s for s in sections if s.name.lower() == section_name.lower()),
                None,
            )
            if not section:
                section = notebook.create_section(name=section_name)
            page_html = (
                f'<!DOCTYPE html><html><head><title>{title}</title></head>'
                f'<body>{html_content}</body></html>'
            )
            page = section.create_page(content=page_html)
            return f"OneNote page '{title}' created" if page else "Page creation failed"
        except Exception as exc:
            return f"OneNote error: {exc}"

    # ── Excel ─────────────────────────────────────────────────────────────────

    def excel_update(self, onedrive_path, sheet_name, cell_range, values):
        """
        Update cells in an Excel workbook stored on OneDrive via the Graph API.
        Values is a flat list; it is reshaped into a 2-D array matching the range.
        """
        drive = self._get_drive()
        item = drive.get_item_by_path(onedrive_path)
        drive_id = drive.object_id
        item_id = item.object_id
        rows, cols = _parse_excel_range(cell_range)
        flat = list(values)
        values_2d = []
        idx = 0
        for _ in range(rows):
            row = [flat[idx + c] if (idx + c) < len(flat) else None for c in range(cols)]
            idx += cols
            values_2d.append(row)
        url = (
            f"https://graph.microsoft.com/v1.0"
            f"/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{sheet_name}')/range(address='{cell_range}')"
        )
        self._graph_patch(url, {"values": values_2d})
        return f"Excel updated: {onedrive_path} [{sheet_name}!{cell_range}]"

    # ── Word ──────────────────────────────────────────────────────────────────

    def word_update(self, onedrive_path, replacements):
        """Download a .docx, replace placeholder text, re-upload to OneDrive."""
        try:
            from docx import Document
        except ImportError:
            return "python-docx not installed. Run: pip install python-docx"
        with tempfile.TemporaryDirectory() as tmpdir:
            local = os.path.join(tmpdir, Path(onedrive_path).name)
            self.onedrive_download(onedrive_path, local)
            doc = Document(local)
            for para in doc.paragraphs:
                for run in para.runs:
                    for old, new in replacements.items():
                        run.text = run.text.replace(old, new)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                for old, new in replacements.items():
                                    run.text = run.text.replace(old, new)
            doc.save(local)
            self.onedrive_upload(local, onedrive_path)
        return f"Word document updated: {onedrive_path}"

    # ── PowerPoint ────────────────────────────────────────────────────────────

    def ppt_update(self, onedrive_path, slide_number, replacements):
        """Download a .pptx, replace text in slide *slide_number* (0-indexed), re-upload."""
        try:
            from pptx import Presentation
        except ImportError:
            return "python-pptx not installed. Run: pip install python-pptx"
        with tempfile.TemporaryDirectory() as tmpdir:
            local = os.path.join(tmpdir, Path(onedrive_path).name)
            self.onedrive_download(onedrive_path, local)
            prs = Presentation(local)
            slide = prs.slides[int(slide_number)]
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            for old, new in replacements.items():
                                run.text = run.text.replace(old, new)
            prs.save(local)
            self.onedrive_upload(local, onedrive_path)
        return f"PowerPoint updated: {onedrive_path} (slide {slide_number})"
PYEOF

cat > "${SRC_DIR}/m365_openclaw/__main__.py" << 'PYEOF'
"""
__main__.py – CLI entry point for the OpenClaw M365 skill.

Invoked as:  python -m m365_openclaw <command> [arguments...]
Or via the  m365  wrapper script placed in ~/.local/bin.
"""

import sys

USAGE = """\
OpenClaw M365 CLI – Microsoft 365 for agents

Commands:
  mail-list [limit]
      List recent inbox messages (default: 20)

  send-mail <to> <subject> <body>
      Send an email to any recipient

  calendar-list [days]
      List upcoming events (default: 7 days)

  calendar-create <subject> <start_iso> <end_iso> [body]
      Create a calendar event (ISO 8601, e.g. 2026-04-01T10:00:00)

  contacts-list [limit]
      List address-book contacts (default: 100)

  contacts-create <first> <last> [email] [phone]
      Add a new contact

  onedrive-list [folder]
      List files in a OneDrive folder (default: /)

  upload <local_path> <remote_path>
      Upload a local file to OneDrive

  download <remote_path> <local_path>
      Download a file from OneDrive

  onenote-create <notebook> <section> <title> <html>
      Create a OneNote page (requires Notes.ReadWrite.All delegated permission)

  excel-update <remote_path> <sheet> <range> [value ...]
      Update Excel cells, e.g.: excel-update report.xlsx Sheet1 A1:B2 Jan Feb 100 200

  word-update <remote_path> key=value [key=value ...]
      Replace placeholders in a Word document, e.g.: word-update doc.docx {{Name}}=Alice

  ppt-update <remote_path> <slide_number> key=value [key=value ...]
      Replace text in a PowerPoint slide (0-indexed), e.g.: ppt-update deck.pptx 0 {{Title}}=Hello
"""


def main():
    if len(sys.argv) < 2:
        print("OpenClaw M365 CLI ready.")
        print(USAGE)
        sys.exit(0)

    cmd = sys.argv[1]
    args = sys.argv[2:]

    try:
        from m365_openclaw.core import M365Client
        client = M365Client()
    except EnvironmentError as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        sys.exit(1)

    try:
        if cmd == "mail-list":
            limit = int(args[0]) if args else 20
            msgs = client.list_mail(limit=limit)
            if not msgs:
                print("No messages found.")
            else:
                for m in msgs:
                    tag = "" if m['is_read'] else "[NEW] "
                    print(f"{tag}{m['date']} | {m['from']}: {m['subject']}")

        elif cmd == "send-mail":
            if len(args) < 3:
                print("Usage: m365 send-mail <to> <subject> <body>")
                sys.exit(1)
            print(client.send_mail(args[0], args[1], args[2]))

        elif cmd == "calendar-list":
            days = int(args[0]) if args else 7
            events = client.get_calendar_events(days=days)
            if not events:
                print(f"No events in the next {days} days.")
            else:
                print(f"Upcoming events (next {days} days):")
                for e in events:
                    loc = f" @ {e['location']}" if e.get('location') else ""
                    print(f"- {e['start']} | {e['subject']}{loc}")

        elif cmd == "calendar-create":
            if len(args) < 3:
                print("Usage: m365 calendar-create <subject> <start_iso> <end_iso> [body]")
                sys.exit(1)
            body = args[3] if len(args) > 3 else ""
            print(client.create_calendar_event(args[0], args[1], args[2], body))

        elif cmd == "contacts-list":
            limit = int(args[0]) if args else 100
            contacts = client.list_contacts(limit=limit)
            if not contacts:
                print("No contacts found.")
            else:
                for c in contacts:
                    emails = ", ".join(c['emails']) if c['emails'] else "(no email)"
                    print(f"- {c['name']} | {emails}")

        elif cmd == "contacts-create":
            if len(args) < 2:
                print("Usage: m365 contacts-create <first> <last> [email] [phone]")
                sys.exit(1)
            email = args[2] if len(args) > 2 else None
            phone = args[3] if len(args) > 3 else None
            print(client.create_contact(args[0], args[1], email, phone))

        elif cmd == "onedrive-list":
            folder = args[0] if args else "/"
            items = client.onedrive_list(folder)
            if not items:
                print("Empty folder.")
            else:
                for item in items:
                    size = f" ({item['size']} B)" if item.get('size') else ""
                    print(f"[{item['type'].upper()}] {item['name']}{size}")

        elif cmd == "upload":
            if len(args) < 2:
                print("Usage: m365 upload <local_path> <remote_path>")
                sys.exit(1)
            print(client.onedrive_upload(args[0], args[1]))

        elif cmd == "download":
            if len(args) < 2:
                print("Usage: m365 download <remote_path> <local_path>")
                sys.exit(1)
            print(client.onedrive_download(args[0], args[1]))

        elif cmd == "onenote-create":
            if len(args) < 4:
                print("Usage: m365 onenote-create <notebook> <section> <title> <html>")
                sys.exit(1)
            print(client.onenote_create_page(args[0], args[1], args[2], args[3]))

        elif cmd == "excel-update":
            if len(args) < 3:
                print("Usage: m365 excel-update <path> <sheet> <range> [value ...]")
                sys.exit(1)
            values = args[3:] if len(args) > 3 else []
            print(client.excel_update(args[0], args[1], args[2], values))

        elif cmd == "word-update":
            if len(args) < 2:
                print("Usage: m365 word-update <path> key=value [key=value ...]")
                sys.exit(1)
            replacements = {}
            for pair in args[1:]:
                if '=' in pair:
                    k, v = pair.split('=', 1)
                    replacements[k] = v
            print(client.word_update(args[0], replacements))

        elif cmd == "ppt-update":
            if len(args) < 3:
                print("Usage: m365 ppt-update <path> <slide_number> key=value [...]")
                sys.exit(1)
            replacements = {}
            for pair in args[2:]:
                if '=' in pair:
                    k, v = pair.split('=', 1)
                    replacements[k] = v
            print(client.ppt_update(args[0], args[1], replacements))

        else:
            print(f"Unknown command: {cmd}", file=sys.stderr)
            print(USAGE)
            sys.exit(1)

    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
PYEOF

# 9. Create SKILL.md for OpenClaw agent auto-discovery
cat > "${INSTALL_DIR}/SKILL.md" << EOF
# SKILL: m365-graph

OpenClaw skill for unattended Microsoft 365 access via Graph API.

## Executable
m365

## Description
Enables OpenClaw agents to send/read mail, manage calendar events, manage
contacts, access OneDrive files, create OneNote pages, and edit Excel,
Word and PowerPoint documents.  Uses client-credentials (daemon) flow.

## Configuration
Credentials stored in: ${CONFIG_FILE}

## Commands

| Command | Description |
|---------|-------------|
| mail-list [N] | List N recent inbox messages (default 20) |
| send-mail <to> <subject> <body> | Send an email |
| calendar-list [days] | List upcoming events (default 7 days) |
| calendar-create <subj> <start> <end> [body] | Create a calendar event (ISO 8601) |
| contacts-list [N] | List N address-book contacts (default 100) |
| contacts-create <first> <last> [email] [phone] | Add a new contact |
| onedrive-list [folder] | List a OneDrive folder (default /) |
| upload <local> <remote> | Upload a file to OneDrive |
| download <remote> <local> | Download a file from OneDrive |
| onenote-create <nb> <sec> <title> <html> | Create a OneNote page |
| excel-update <path> <sheet> <range> [val ...] | Update Excel cells |
| word-update <path> key=val ... | Replace placeholders in a Word document |
| ppt-update <path> <slide> key=val ... | Replace text in a PowerPoint slide |

## Notes
- OneNote requires the delegated permission Notes.ReadWrite.All
- Excel update uses the Graph API directly; the workbook must not be open in Office
- Word and PowerPoint use a download → edit locally → re-upload workflow
EOF

# 10. Start initial authentication
echo ""
echo "[10/10] Starting initial authentication..."
echo "Please sign in with: ${M365_USER}"
echo "If browser does not open (headless system), copy the URL from terminal."
echo ""

export PYTHONPATH="${SRC_DIR}:${PYTHONPATH:-}"
"${VENV_DIR}/bin/python" -m m365_openclaw calendar-list

echo ""
echo "=== Installation complete! ==="
echo ""
echo "Available commands:"
echo "  m365 mail-list              – list recent inbox messages"
echo "  m365 send-mail <to> <sub> <body>  – send an email"
echo "  m365 calendar-list          – list upcoming calendar events"
echo "  m365 calendar-create ...    – create a calendar event"
echo "  m365 contacts-list          – list contacts"
echo "  m365 contacts-create ...    – add a contact"
echo "  m365 onedrive-list [folder] – list OneDrive folder"
echo "  m365 upload <local> <remote>  – upload file to OneDrive"
echo "  m365 download <remote> <local> – download file from OneDrive"
echo "  m365 onenote-create ...     – create a OneNote page"
echo "  m365 excel-update ...       – update Excel cells"
echo "  m365 word-update ...        – replace text in Word document"
echo "  m365 ppt-update ...         – replace text in PowerPoint slide"
echo ""
echo "If authentication fails or no calendar appears:"
echo "  rm -rf ${OPENCLAW_DIR}/credentials/m365_token_cache.bin"
echo "  cd ${INSTALL_DIR}"
echo "  ${VENV_DIR}/bin/python -m m365_openclaw calendar-list"
echo ""
echo "Done!"
