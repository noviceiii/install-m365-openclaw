#!/usr/bin/env bash
# =============================================================================
# OpenClaw M365 Graph Skill - Installer 
# =============================================================================
# Features:
# - Asks for OpenClaw base directory and M365 user email
# - Creates proper package structure: src/m365_openclaw/__init__.py
# - Installs as editable package with correct setup.py
# - Full M365Client: Mail, Calendar, Contacts, OneDrive, OneNote, Excel, Word, PPT, ToDo
# - All CLI commands implemented in __main__.py (mail, calendar, contacts, OneDrive, ToDo)
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
    python-docx python-pptx openpyxl requests

# 4. Create proper package structure and setup.py
touch "${SRC_DIR}/m365_openclaw/__init__.py"

cat > "${INSTALL_DIR}/setup.py" << 'EOF'
from setuptools import setup, find_packages

setup(
    name='m365_openclaw',
    version='0.2.0',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'O365',
        'msal',
        'msal_extensions',
        'python-dotenv',
        'python-docx',
        'python-pptx',
        'openpyxl',
        'requests',
    ],
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
echo "                Files.ReadWrite.All, Contacts.ReadWrite, Tasks.ReadWrite"
echo "   Delegated:   Notes.ReadWrite.All, offline_access"
echo "   → Grant admin consent"
echo "3. Certificates & secrets → New client secret → copy Value"
echo ""
echo "IMPORTANT – Mail access (Exchange Online Application Access Policy):"
echo "   Mail.ReadWrite.All alone may not be sufficient for daemon/app-only access."
echo "   An Exchange admin must run this in Exchange Online PowerShell to permit"
echo "   the app to access the target mailbox:"
echo ""
echo "     New-ApplicationAccessPolicy \\"
echo "       -AppId <CLIENT_ID> \\"
echo "       -PolicyScopeGroupId <M365_USER_EMAIL> \\"
echo "       -AccessRight RestrictAccess \\"
echo "       -Description 'OpenClaw M365 mail access'"
echo ""
echo "   Reference: https://aka.ms/graph-app-access-policy"
echo "   (Calendar and Contacts do not require this extra policy.)"
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

Supports: Mail, Calendar, Contacts, OneDrive, OneNote, Excel, Word, PowerPoint,
          ToDo Tasks.
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

import requests
from dotenv import load_dotenv
from O365 import Account
from O365.utils import FileSystemTokenBackend

load_dotenv()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


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

    def _auth_headers(self, content_type="application/json"):
        headers = {'Authorization': f'Bearer {self._access_token()}'}
        if content_type:
            headers['Content-Type'] = content_type
        return headers

    def _graph_get(self, url):
        """HTTP GET to Graph API."""
        resp = requests.get(url, headers=self._auth_headers(content_type=None), timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_post(self, url, data=None, json_data=None):
        """HTTP POST to Graph API."""
        payload = json_data if json_data is not None else data
        resp = requests.post(url, headers=self._auth_headers(), json=payload, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_patch(self, url, data):
        """HTTP PATCH to the Microsoft Graph API with JSON body."""
        resp = requests.patch(url, headers=self._auth_headers(), json=data, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_delete(self, url):
        """HTTP DELETE to Graph API."""
        resp = requests.delete(url, headers=self._auth_headers(content_type=None), timeout=30)
        resp.raise_for_status()
        return {}

    def _graph_put(self, url, data, content_type):
        """HTTP PUT to Graph API (used for binary uploads)."""
        headers = {
            'Authorization': f'Bearer {self._access_token()}',
            'Content-Type': content_type,
        }
        resp = requests.put(url, headers=headers, data=data, timeout=60)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    # ── Mail ──────────────────────────────────────────────────────────────────

    @staticmethod
    def _raise_if_mail_403(exc):
        """Re-raise with actionable guidance when a 403 occurs on a mail endpoint."""
        if isinstance(exc, requests.exceptions.HTTPError):
            resp = getattr(exc, 'response', None)
            if resp is not None and resp.status_code == 403:
                url = getattr(resp, 'url', '') or ''
                if 'mailFolders' in url or '/messages' in url:
                    raise PermissionError(
                        "Mail access denied (HTTP 403 Forbidden).\n"
                        "\n"
                        "Common causes and fixes:\n"
                        "  1. The 'Mail.ReadWrite.All' application permission is not\n"
                        "     admin-consented in your Azure App Registration.\n"
                        "     Go to: Entra ID → App registrations → <your app>\n"
                        "             → API permissions → Grant admin consent\n"
                        "\n"
                        "  2. Exchange Online requires an Application Access Policy for\n"
                        "     mail access via application credentials (daemon/app flow).\n"
                        "     Ask your Exchange admin to run this in Exchange Online PowerShell:\n"
                        "\n"
                        "       New-ApplicationAccessPolicy \\\n"
                        "         -AppId <CLIENT_ID> \\\n"
                        "         -PolicyScopeGroupId <M365_USER_EMAIL> \\\n"
                        "         -AccessRight RestrictAccess \\\n"
                        "         -Description 'OpenClaw M365 mail access'\n"
                        "\n"
                        "     Reference: https://aka.ms/graph-app-access-policy"
                    ) from exc

    def send_mail(self, to_address, subject, body, cc=None, bcc=None,
                  sensitivity="Normal", importance="Normal", attachments=None,
                  request_delivery_receipt=False, request_read_receipt=False):
        """Send an email with optional CC, BCC, attachments, and receipt requests."""
        try:
            m = self.account.mailbox().new_message()
            if isinstance(to_address, list):
                for addr in to_address:
                    m.to.add(addr)
            else:
                m.to.add(to_address)
            m.subject = subject
            m.body = body
            if cc:
                for addr in ([cc] if isinstance(cc, str) else cc):
                    m.cc.add(addr)
            if bcc:
                for addr in ([bcc] if isinstance(bcc, str) else bcc):
                    m.bcc.add(addr)
            m.sensitivity = sensitivity
            m.importance = importance
            if attachments:
                for path in attachments:
                    m.attachments.add(path)
            if request_delivery_receipt:
                m.request_delivery_receipt = True
            if request_read_receipt:
                m.request_read_receipt = True
            m.send()
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        to_str = to_address if isinstance(to_address, str) else ", ".join(to_address)
        return f"Email sent to {to_str}"

    def list_mail(self, limit=20):
        """Return recent messages from the inbox."""
        try:
            messages = self.account.mailbox().inbox_folder().get_messages(limit=limit)
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        result = []
        for msg in messages:
            result.append({
                'id': msg.object_id,
                'subject': msg.subject or '(no subject)',
                'from': str(msg.sender),
                'date': msg.received.isoformat() if msg.received else None,
                'is_read': msg.is_read,
            })
        return result

    def mail_list(self, limit=20, folder="inbox"):
        """List messages from a folder: 'inbox', 'sent', 'drafts', 'deleted'."""
        folder_map = {
            "inbox": "inbox",
            "sent": "sentitems",
            "drafts": "drafts",
            "deleted": "deleteditems",
        }
        folder_id = folder_map.get(folder.lower(), folder)
        url = f"{GRAPH_BASE}/users/{self.user}/mailFolders/{folder_id}/messages?$top={limit}&$orderby=receivedDateTime desc"
        try:
            data = self._graph_get(url)
        except requests.exceptions.HTTPError as exc:
            self._raise_if_mail_403(exc)
            raise
        result = []
        for msg in data.get('value', []):
            sender = msg.get('from', {}).get('emailAddress', {})
            result.append({
                'id': msg.get('id'),
                'subject': msg.get('subject') or '(no subject)',
                'from': sender.get('address', ''),
                'date': msg.get('receivedDateTime'),
                'is_read': msg.get('isRead', True),
            })
        return result

    def mail_read_by_subject(self, subject, limit=5):
        """Find and return messages matching the subject. Returns list of dicts."""
        url = (
            f"{GRAPH_BASE}/users/{self.user}/messages"
            f"?$filter=contains(subject,'{subject}')"
            f"&$top={limit}&$orderby=receivedDateTime desc"
        )
        try:
            data = self._graph_get(url)
        except requests.exceptions.HTTPError as exc:
            self._raise_if_mail_403(exc)
            raise
        result = []
        for msg in data.get('value', []):
            sender = msg.get('from', {}).get('emailAddress', {})
            result.append({
                'id': msg.get('id'),
                'subject': msg.get('subject') or '(no subject)',
                'from': sender.get('address', ''),
                'date': msg.get('receivedDateTime'),
                'body': msg.get('body', {}).get('content', ''),
            })
        return result

    def mail_read_headers(self, message_id):
        """Return internet message headers of a specific message."""
        url = f"{GRAPH_BASE}/users/{self.user}/messages/{message_id}?$select=internetMessageHeaders,subject,from,receivedDateTime"
        try:
            return self._graph_get(url)
        except requests.exceptions.HTTPError as exc:
            self._raise_if_mail_403(exc)
            raise

    def mail_reply(self, message_id, body):
        """Reply to the sender of a message."""
        url = f"{GRAPH_BASE}/users/{self.user}/messages/{message_id}/reply"
        try:
            self._graph_post(url, json_data={"comment": body})
        except requests.exceptions.HTTPError as exc:
            self._raise_if_mail_403(exc)
            raise
        return "Reply sent"

    def mail_reply_all(self, message_id, body):
        """Reply to all recipients of a message."""
        url = f"{GRAPH_BASE}/users/{self.user}/messages/{message_id}/replyAll"
        try:
            self._graph_post(url, json_data={"comment": body})
        except requests.exceptions.HTTPError as exc:
            self._raise_if_mail_403(exc)
            raise
        return "Reply-all sent"

    def mail_forward(self, message_id, to_address, body=""):
        """Forward a message to a new recipient."""
        url = f"{GRAPH_BASE}/users/{self.user}/messages/{message_id}/forward"
        payload = {
            "comment": body,
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        }
        try:
            self._graph_post(url, json_data=payload)
        except requests.exceptions.HTTPError as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Message forwarded to {to_address}"

    def mail_delete(self, message_id):
        """Delete a message by ID."""
        url = f"{GRAPH_BASE}/users/{self.user}/messages/{message_id}"
        try:
            self._graph_delete(url)
        except requests.exceptions.HTTPError as exc:
            self._raise_if_mail_403(exc)
            raise
        return "Message deleted"

    def mail_move(self, message_id, folder_name):
        """Move a message to a named folder."""
        url = f"{GRAPH_BASE}/users/{self.user}/messages/{message_id}/move"
        try:
            self._graph_post(url, json_data={"destinationId": folder_name})
        except requests.exceptions.HTTPError as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Message moved to {folder_name}"

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

    def create_calendar_event(self, subject, start_iso, end_iso, body="", location="",
                              required_attendees=None, optional_attendees=None,
                              is_private=False, reminder_minutes=15, attachments=None):
        """Create a calendar event with attendees, privacy, and reminder support."""
        schedule = self.account.schedule()
        calendar = schedule.get_default_calendar()
        event = calendar.new_event()
        event.subject = subject
        event.body = body
        if location:
            event.location = location
        event.start = datetime.fromisoformat(start_iso)
        event.end = datetime.fromisoformat(end_iso)
        if required_attendees:
            for email in required_attendees:
                event.attendees.add(email, attendee_type='required')
        if optional_attendees:
            for email in optional_attendees:
                event.attendees.add(email, attendee_type='optional')
        if is_private:
            event.sensitivity = 'private'
        event.remind_before_minutes = reminder_minutes
        if attachments:
            for path in attachments:
                event.attachments.add(path)
        event.save()
        return f"Event '{subject}' created"

    # ── Contacts ──────────────────────────────────────────────────────────────

    def list_contacts(self, limit=100):
        """List contacts: first name, last name, biz email, personal email, phones."""
        address_book = self.account.address_book()
        contacts = address_book.get_contacts(limit=limit)
        result = []
        for contact in contacts:
            emails_raw = list(contact.emails) if hasattr(contact, 'emails') and contact.emails else []
            email_biz = None
            email_personal = None
            for e in emails_raw:
                e_str = str(e)
                addr_type = getattr(e, 'email_type', '') or ''
                if 'work' in addr_type.lower() or 'business' in addr_type.lower():
                    email_biz = email_biz or e_str
                elif 'home' in addr_type.lower() or 'personal' in addr_type.lower():
                    email_personal = email_personal or e_str
                else:
                    email_biz = email_biz or e_str

            biz_phones = list(contact.business_phones) if hasattr(contact, 'business_phones') and contact.business_phones else []
            home_phones = list(contact.home_phones) if hasattr(contact, 'home_phones') and contact.home_phones else []
            mobile_phones = list(contact.mobile_phone) if hasattr(contact, 'mobile_phone') and contact.mobile_phone else []

            result.append({
                'id': contact.object_id,
                'given_name': contact.given_name or '',
                'surname': contact.surname or '',
                'email_business': email_biz,
                'email_personal': email_personal,
                'phone_business': biz_phones[0] if biz_phones else None,
                'phone_mobile': mobile_phones[0] if mobile_phones else None,
                'phone_home': home_phones[0] if home_phones else None,
            })
        return result

    def get_contact(self, contact_id):
        """Get all fields of a specific contact by ID."""
        url = f"{GRAPH_BASE}/users/{self.user}/contacts/{contact_id}"
        return self._graph_get(url)

    def create_contact(self, given_name, surname, email=None, phone=None,
                       email_business=None, email_personal=None,
                       phone_business=None, phone_mobile=None, phone_home=None,
                       street_business=None, city_business=None, zip_business=None, country_business=None,
                       street_personal=None, city_personal=None, zip_personal=None, country_personal=None,
                       birthday=None, anniversary=None, website_business=None, website_personal=None,
                       spouse=None, notes=None):
        """Create a new contact. Legacy positional args email/phone map to biz email/phone."""
        email_business = email_business or email
        phone_business = phone_business or phone

        address_book = self.account.address_book()
        contact = address_book.new_contact()
        contact.given_name = given_name
        contact.surname = surname
        if email_business:
            contact.emails.add(email_business, email_type='business')
        if email_personal:
            contact.emails.add(email_personal, email_type='personal')
        if phone_business:
            contact.business_phones.append(phone_business)
        if phone_mobile:
            contact.mobile_phone = phone_mobile
        if phone_home:
            contact.home_phones.append(phone_home)
        if birthday:
            try:
                contact.birthday = datetime.fromisoformat(birthday)
            except ValueError:
                pass
        if notes:
            contact.personal_notes = notes
        contact.save()
        return f"Contact '{given_name} {surname}' created"

    def update_contact_photo(self, contact_id, photo_path):
        """Upload a photo for a contact."""
        url = f"{GRAPH_BASE}/users/{self.user}/contacts/{contact_id}/photo/$value"
        ext = Path(photo_path).suffix.lower()
        content_type = "image/jpeg" if ext in ('.jpg', '.jpeg') else "image/png"
        with open(photo_path, 'rb') as f:
            data = f.read()
        self._graph_put(url, data, content_type)
        return f"Photo updated for contact {contact_id}"

    def delete_contact_photo(self, contact_id):
        """Delete a contact's photo."""
        url = f"{GRAPH_BASE}/users/{self.user}/contacts/{contact_id}/photo/$value"
        self._graph_delete(url)
        return f"Photo deleted for contact {contact_id}"

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
        Requires the delegated permission Notes.ReadWrite.All.
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
            f"{GRAPH_BASE}"
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

    # ── ToDo / Tasks ──────────────────────────────────────────────────────────

    def todo_list_lists(self):
        """Return all To Do task lists."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists"
        data = self._graph_get(url)
        return data.get('value', [])

    def todo_create_list(self, name):
        """Create a new task list."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists"
        result = self._graph_post(url, json_data={"displayName": name})
        return result

    def todo_rename_list(self, list_id, new_name):
        """Rename a task list."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}"
        return self._graph_patch(url, {"displayName": new_name})

    def todo_delete_list(self, list_id):
        """Delete a task list."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}"
        self._graph_delete(url)
        return f"Task list {list_id} deleted"

    def todo_list_tasks(self, list_id, due_before=None, due_after=None):
        """List all tasks in a list, optionally filtered by due date (YYYY-MM-DD)."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}/tasks"
        data = self._graph_get(url)
        tasks = data.get('value', [])
        if due_before:
            cutoff = datetime.fromisoformat(due_before).replace(tzinfo=timezone.utc)
            tasks = [t for t in tasks if self._task_due(t) and self._task_due(t) <= cutoff]
        if due_after:
            floor = datetime.fromisoformat(due_after).replace(tzinfo=timezone.utc)
            tasks = [t for t in tasks if self._task_due(t) and self._task_due(t) >= floor]
        return tasks

    def todo_list_all_tasks(self, due_before=None, due_after=None):
        """List all tasks across all lists."""
        lists = self.todo_list_lists()
        all_tasks = []
        for lst in lists:
            tasks = self.todo_list_tasks(lst['id'], due_before=due_before, due_after=due_after)
            for t in tasks:
                t['_list_id'] = lst['id']
                t['_list_name'] = lst.get('displayName', '')
            all_tasks.extend(tasks)
        return all_tasks

    @staticmethod
    def _task_due(task):
        """Parse due datetime from a task dict, returns datetime or None."""
        due = task.get('dueDateTime')
        if not due:
            return None
        try:
            dt = datetime.fromisoformat(due.get('dateTime', '').rstrip('Z'))
            return dt.replace(tzinfo=timezone.utc)
        except (ValueError, AttributeError):
            return None

    def todo_create_task(self, list_id, title, note=None, due_date=None, reminder_datetime=None):
        """Create a task in a list."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}/tasks"
        payload = {"title": title}
        if note:
            payload["body"] = {"content": note, "contentType": "text"}
        if due_date:
            payload["dueDateTime"] = {"dateTime": f"{due_date}T00:00:00", "timeZone": "UTC"}
        if reminder_datetime:
            payload["reminderDateTime"] = {"dateTime": reminder_datetime, "timeZone": "UTC"}
        return self._graph_post(url, json_data=payload)

    def todo_update_task(self, list_id, task_id, title=None, note=None, due_date=None, reminder_datetime=None):
        """Update a task's fields."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}/tasks/{task_id}"
        payload = {}
        if title is not None:
            payload["title"] = title
        if note is not None:
            payload["body"] = {"content": note, "contentType": "text"}
        if due_date is not None:
            payload["dueDateTime"] = {"dateTime": f"{due_date}T00:00:00", "timeZone": "UTC"}
        if reminder_datetime is not None:
            payload["reminderDateTime"] = {"dateTime": reminder_datetime, "timeZone": "UTC"}
        return self._graph_patch(url, payload)

    def todo_complete_task(self, list_id, task_id):
        """Mark a task as completed."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}/tasks/{task_id}"
        return self._graph_patch(url, {"status": "completed"})

    def todo_delete_task(self, list_id, task_id):
        """Delete a task by ID."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}/tasks/{task_id}"
        self._graph_delete(url)
        return f"Task {task_id} deleted"

    def todo_add_step(self, list_id, task_id, title):
        """Add a checklist step to a task."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}/tasks/{task_id}/checklistItems"
        return self._graph_post(url, json_data={"displayName": title})

    def todo_complete_step(self, list_id, task_id, step_id):
        """Mark a checklist step as completed."""
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}/tasks/{task_id}/checklistItems/{step_id}"
        return self._graph_patch(url, {"isChecked": True})

    def todo_assign_task(self, list_id, task_id, assignee_email):
        """
        Note: MS To Do API has limited assignment support; this stores the
        assignee email in the task body as a workaround.
        """
        url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{list_id}/tasks/{task_id}"
        task = self._graph_get(url)
        existing_body = task.get('body', {}).get('content', '')
        new_body = f"Assigned to: {assignee_email}\n{existing_body}"
        return self._graph_patch(url, {"body": {"content": new_body, "contentType": "text"}})

    def todo_move_task(self, source_list_id, task_id, target_list_id):
        """Move a task from one list to another by recreating it in the target."""
        src_url = f"{GRAPH_BASE}/users/{self.user}/todo/lists/{source_list_id}/tasks/{task_id}"
        task = self._graph_get(src_url)
        payload = {k: task[k] for k in ('title', 'body', 'dueDateTime', 'reminderDateTime', 'status')
                   if k in task}
        created = self.todo_create_task(
            target_list_id,
            payload.get('title', ''),
            note=payload.get('body', {}).get('content'),
            due_date=payload.get('dueDateTime', {}).get('dateTime', '').split('T')[0] or None,
        )
        self._graph_delete(src_url)
        return created
PYEOF

cat > "${SRC_DIR}/m365_openclaw/__main__.py" << 'PYEOF'
"""
__main__.py – CLI entry point for the OpenClaw M365 skill.

Invoked as:  python -m m365_openclaw <command> [arguments...]
Or via the  m365  wrapper script placed in ~/.local/bin.
"""

import argparse
import json
import sys

USAGE = """\
OpenClaw M365 CLI – Microsoft 365 for agents

Mail commands:
  mail-list [limit]
      List recent inbox messages (default: 20)

  send-mail <to> <subject> <body> [options]
      Send an email. Options: --cc, --bcc, --importance, --sensitivity,
      --attach, --delivery-receipt, --read-receipt

  mail-read-subject <subject>
      Find messages matching a subject

  mail-read <message-id>
      Read a specific message (returns headers + body)

  mail-reply <message-id> <body>
      Reply to the sender of a message

  mail-reply-all <message-id> <body>
      Reply to all recipients of a message

  mail-forward <message-id> <to> [body]
      Forward a message to a new recipient

  mail-delete <message-id>
      Delete a message by ID

  mail-move <message-id> <folder>
      Move a message to a named folder

Calendar commands:
  calendar-list [days]
      List upcoming events (default: 7 days)

  calendar-create <subject> <start_iso> <end_iso> [body] [options]
      Create a calendar event (ISO 8601, e.g. 2026-04-01T10:00:00)
      Options: --location, --required, --optional, --private,
               --reminder-minutes, --attach

Contacts commands:
  contacts-list [limit]
      List address-book contacts (default: 100)

  contacts-get <contact-id>
      Get all fields of a contact

  contacts-create <first> <last> [options]
      Add a new contact. Options: --email-business, --email-personal,
      --phone-business, --phone-mobile, --phone-home, --birthday, --notes

  contacts-photo-update <contact-id> <photo-path>
      Upload a photo for a contact

  contacts-photo-delete <contact-id>
      Delete a contact's photo

OneDrive commands:
  onedrive-list [folder]
      List files in a OneDrive folder (default: /)

  upload <local_path> <remote_path>
      Upload a local file to OneDrive

  download <remote_path> <local_path>
      Download a file from OneDrive

Document commands:
  onenote-create <notebook> <section> <title> <html>
      Create a OneNote page (requires Notes.ReadWrite.All delegated permission)

  excel-update <remote_path> <sheet> <range> [value ...]
      Update Excel cells

  word-update <remote_path> key=value [key=value ...]
      Replace placeholders in a Word document

  ppt-update <remote_path> <slide_number> key=value [key=value ...]
      Replace text in a PowerPoint slide (0-indexed)

ToDo commands:
  todo-lists
      List all task lists

  todo-list-create <name>
      Create a task list

  todo-list-rename <list-id> <new-name>
      Rename a task list

  todo-list-delete <list-id>
      Delete a task list

  todo-tasks <list-id> [--due-before DATE] [--due-after DATE]
      List tasks in a list

  todo-tasks-all [--due-before DATE] [--due-after DATE]
      List all tasks across all lists

  todo-task-create <list-id> <title> [--note text] [--due YYYY-MM-DD] [--reminder datetime]
      Create a task

  todo-task-update <list-id> <task-id> [--title t] [--note n] [--due d] [--reminder dt]
      Update a task

  todo-task-complete <list-id> <task-id>
      Mark a task as completed

  todo-task-delete <list-id> <task-id>
      Delete a task

  todo-step-add <list-id> <task-id> <title>
      Add a step (checklist item) to a task

  todo-step-complete <list-id> <task-id> <step-id>
      Mark a step as completed

  todo-task-assign <list-id> <task-id> <email>
      Assign a task to a person

  todo-task-move <source-list-id> <task-id> <target-list-id>
      Move a task between lists
"""


def _print_json(data):
    print(json.dumps(data, indent=2, default=str))


def _parse_send_mail_args(args):
    """Parse extended send-mail arguments."""
    parser = argparse.ArgumentParser(prog="m365 send-mail", add_help=False)
    parser.add_argument("to")
    parser.add_argument("subject")
    parser.add_argument("body")
    parser.add_argument("--cc", action="append", default=[])
    parser.add_argument("--bcc", action="append", default=[])
    parser.add_argument("--importance", default="Normal",
                        choices=["High", "Normal", "Low"])
    parser.add_argument("--sensitivity", default="Normal",
                        choices=["Normal", "Personal", "Private", "Confidential"])
    parser.add_argument("--attach", action="append", default=[], dest="attachments")
    parser.add_argument("--delivery-receipt", action="store_true")
    parser.add_argument("--read-receipt", action="store_true")
    return parser.parse_args(args)


def _parse_calendar_create_args(args):
    """Parse extended calendar-create arguments."""
    parser = argparse.ArgumentParser(prog="m365 calendar-create", add_help=False)
    parser.add_argument("subject")
    parser.add_argument("start_iso")
    parser.add_argument("end_iso")
    parser.add_argument("body", nargs="?", default="")
    parser.add_argument("--location", default="")
    parser.add_argument("--required", action="append", default=[], dest="required_attendees")
    parser.add_argument("--optional", action="append", default=[], dest="optional_attendees")
    parser.add_argument("--private", action="store_true")
    parser.add_argument("--reminder-minutes", type=int, default=15)
    parser.add_argument("--attach", action="append", default=[], dest="attachments")
    return parser.parse_args(args)


def _parse_contacts_create_args(args):
    """Parse extended contacts-create arguments."""
    parser = argparse.ArgumentParser(prog="m365 contacts-create", add_help=False)
    parser.add_argument("first")
    parser.add_argument("last")
    parser.add_argument("email_pos", nargs="?", default=None, metavar="email")
    parser.add_argument("phone_pos", nargs="?", default=None, metavar="phone")
    parser.add_argument("--email-business", default=None)
    parser.add_argument("--email-personal", default=None)
    parser.add_argument("--phone-business", default=None)
    parser.add_argument("--phone-mobile", default=None)
    parser.add_argument("--phone-home", default=None)
    parser.add_argument("--birthday", default=None)
    parser.add_argument("--notes", default=None)
    return parser.parse_args(args)


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
        # ── Mail ──────────────────────────────────────────────────────────────

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
                print("Usage: m365 send-mail <to> <subject> <body> [--cc addr] [--bcc addr] "
                      "[--importance High|Normal|Low] [--sensitivity Normal|...] "
                      "[--attach path] [--delivery-receipt] [--read-receipt]")
                sys.exit(1)
            ns = _parse_send_mail_args(args)
            print(client.send_mail(
                ns.to, ns.subject, ns.body,
                cc=ns.cc or None,
                bcc=ns.bcc or None,
                sensitivity=ns.sensitivity,
                importance=ns.importance,
                attachments=ns.attachments or None,
                request_delivery_receipt=ns.delivery_receipt,
                request_read_receipt=ns.read_receipt,
            ))

        elif cmd == "mail-read-subject":
            if not args:
                print("Usage: m365 mail-read-subject <subject>")
                sys.exit(1)
            results = client.mail_read_by_subject(" ".join(args))
            if not results:
                print("No messages found matching that subject.")
            else:
                for msg in results:
                    print(f"--- {msg['date']} | From: {msg['from']} | Subject: {msg['subject']} ---")
                    print(msg.get('body', ''))
                    print()

        elif cmd == "mail-read":
            if not args:
                print("Usage: m365 mail-read <message-id>")
                sys.exit(1)
            _print_json(client.mail_read_headers(args[0]))

        elif cmd == "mail-reply":
            if len(args) < 2:
                print("Usage: m365 mail-reply <message-id> <body>")
                sys.exit(1)
            print(client.mail_reply(args[0], " ".join(args[1:])))

        elif cmd == "mail-reply-all":
            if len(args) < 2:
                print("Usage: m365 mail-reply-all <message-id> <body>")
                sys.exit(1)
            print(client.mail_reply_all(args[0], " ".join(args[1:])))

        elif cmd == "mail-forward":
            if len(args) < 2:
                print("Usage: m365 mail-forward <message-id> <to> [body]")
                sys.exit(1)
            body = " ".join(args[2:]) if len(args) > 2 else ""
            print(client.mail_forward(args[0], args[1], body))

        elif cmd == "mail-delete":
            if not args:
                print("Usage: m365 mail-delete <message-id>")
                sys.exit(1)
            print(client.mail_delete(args[0]))

        elif cmd == "mail-move":
            if len(args) < 2:
                print("Usage: m365 mail-move <message-id> <folder>")
                sys.exit(1)
            print(client.mail_move(args[0], args[1]))

        # ── Calendar ──────────────────────────────────────────────────────────

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
                print("Usage: m365 calendar-create <subject> <start_iso> <end_iso> [body] "
                      "[--location loc] [--required email] [--optional email] "
                      "[--private] [--reminder-minutes N] [--attach path]")
                sys.exit(1)
            ns = _parse_calendar_create_args(args)
            print(client.create_calendar_event(
                ns.subject, ns.start_iso, ns.end_iso, ns.body,
                location=ns.location,
                required_attendees=ns.required_attendees or None,
                optional_attendees=ns.optional_attendees or None,
                is_private=ns.private,
                reminder_minutes=ns.reminder_minutes,
                attachments=ns.attachments or None,
            ))

        # ── Contacts ──────────────────────────────────────────────────────────

        elif cmd == "contacts-list":
            limit = int(args[0]) if args else 100
            contacts = client.list_contacts(limit=limit)
            if not contacts:
                print("No contacts found.")
            else:
                header = f"{'First':<15} {'Last':<20} {'Biz Email':<35} {'Pers Email':<35} {'Biz Phone':<18} {'Mobile':<18}"
                print(header)
                print("-" * len(header))
                for c in contacts:
                    print(
                        f"{c['given_name']:<15} {c['surname']:<20} "
                        f"{(c['email_business'] or ''):<35} {(c['email_personal'] or ''):<35} "
                        f"{(c['phone_business'] or ''):<18} {(c['phone_mobile'] or ''):<18}"
                    )

        elif cmd == "contacts-get":
            if not args:
                print("Usage: m365 contacts-get <contact-id>")
                sys.exit(1)
            _print_json(client.get_contact(args[0]))

        elif cmd == "contacts-create":
            if len(args) < 2:
                print("Usage: m365 contacts-create <first> <last> [email] [phone] "
                      "[--email-business e] [--email-personal e] [--phone-business p] "
                      "[--phone-mobile p] [--phone-home p] [--birthday YYYY-MM-DD] [--notes text]")
                sys.exit(1)
            ns = _parse_contacts_create_args(args)
            email_biz = ns.email_business or ns.email_pos
            phone_biz = ns.phone_business or ns.phone_pos
            print(client.create_contact(
                ns.first, ns.last,
                email_business=email_biz,
                email_personal=ns.email_personal,
                phone_business=phone_biz,
                phone_mobile=ns.phone_mobile,
                phone_home=ns.phone_home,
                birthday=ns.birthday,
                notes=ns.notes,
            ))

        elif cmd == "contacts-photo-update":
            if len(args) < 2:
                print("Usage: m365 contacts-photo-update <contact-id> <photo-path>")
                sys.exit(1)
            print(client.update_contact_photo(args[0], args[1]))

        elif cmd == "contacts-photo-delete":
            if not args:
                print("Usage: m365 contacts-photo-delete <contact-id>")
                sys.exit(1)
            print(client.delete_contact_photo(args[0]))

        # ── OneDrive ──────────────────────────────────────────────────────────

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

        # ── OneNote ───────────────────────────────────────────────────────────

        elif cmd == "onenote-create":
            if len(args) < 4:
                print("Usage: m365 onenote-create <notebook> <section> <title> <html>")
                sys.exit(1)
            print(client.onenote_create_page(args[0], args[1], args[2], args[3]))

        # ── Excel / Word / PowerPoint ─────────────────────────────────────────

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

        # ── ToDo ──────────────────────────────────────────────────────────────

        elif cmd == "todo-lists":
            lists = client.todo_list_lists()
            if not lists:
                print("No task lists found.")
            else:
                for lst in lists:
                    print(f"{lst.get('id')} | {lst.get('displayName')}")

        elif cmd == "todo-list-create":
            if not args:
                print("Usage: m365 todo-list-create <name>")
                sys.exit(1)
            result = client.todo_create_list(" ".join(args))
            print(f"Created list: {result.get('displayName')} (id: {result.get('id')})")

        elif cmd == "todo-list-rename":
            if len(args) < 2:
                print("Usage: m365 todo-list-rename <list-id> <new-name>")
                sys.exit(1)
            result = client.todo_rename_list(args[0], " ".join(args[1:]))
            print(f"Renamed to: {result.get('displayName')}")

        elif cmd == "todo-list-delete":
            if not args:
                print("Usage: m365 todo-list-delete <list-id>")
                sys.exit(1)
            print(client.todo_delete_list(args[0]))

        elif cmd == "todo-tasks":
            if not args:
                print("Usage: m365 todo-tasks <list-id> [--due-before DATE] [--due-after DATE]")
                sys.exit(1)
            parser = argparse.ArgumentParser(prog="m365 todo-tasks", add_help=False)
            parser.add_argument("list_id")
            parser.add_argument("--due-before", default=None)
            parser.add_argument("--due-after", default=None)
            ns = parser.parse_args(args)
            tasks = client.todo_list_tasks(ns.list_id, due_before=ns.due_before, due_after=ns.due_after)
            if not tasks:
                print("No tasks found.")
            else:
                for t in tasks:
                    due = t.get('dueDateTime', {}).get('dateTime', '')[:10] if t.get('dueDateTime') else ''
                    status = t.get('status', '')
                    print(f"{t.get('id')} | {t.get('title')} | due: {due} | {status}")

        elif cmd == "todo-tasks-all":
            parser = argparse.ArgumentParser(prog="m365 todo-tasks-all", add_help=False)
            parser.add_argument("--due-before", default=None)
            parser.add_argument("--due-after", default=None)
            ns = parser.parse_args(args)
            tasks = client.todo_list_all_tasks(due_before=ns.due_before, due_after=ns.due_after)
            if not tasks:
                print("No tasks found.")
            else:
                for t in tasks:
                    due = t.get('dueDateTime', {}).get('dateTime', '')[:10] if t.get('dueDateTime') else ''
                    status = t.get('status', '')
                    list_name = t.get('_list_name', '')
                    print(f"[{list_name}] {t.get('title')} | due: {due} | {status}")

        elif cmd == "todo-task-create":
            if len(args) < 2:
                print("Usage: m365 todo-task-create <list-id> <title> [--note text] "
                      "[--due YYYY-MM-DD] [--reminder datetime]")
                sys.exit(1)
            parser = argparse.ArgumentParser(prog="m365 todo-task-create", add_help=False)
            parser.add_argument("list_id")
            parser.add_argument("title")
            parser.add_argument("--note", default=None)
            parser.add_argument("--due", default=None)
            parser.add_argument("--reminder", default=None)
            ns = parser.parse_args(args)
            result = client.todo_create_task(ns.list_id, ns.title, note=ns.note,
                                             due_date=ns.due, reminder_datetime=ns.reminder)
            print(f"Task created: {result.get('title')} (id: {result.get('id')})")

        elif cmd == "todo-task-update":
            if len(args) < 2:
                print("Usage: m365 todo-task-update <list-id> <task-id> [--title t] "
                      "[--note n] [--due d] [--reminder dt]")
                sys.exit(1)
            parser = argparse.ArgumentParser(prog="m365 todo-task-update", add_help=False)
            parser.add_argument("list_id")
            parser.add_argument("task_id")
            parser.add_argument("--title", default=None)
            parser.add_argument("--note", default=None)
            parser.add_argument("--due", default=None)
            parser.add_argument("--reminder", default=None)
            ns = parser.parse_args(args)
            result = client.todo_update_task(ns.list_id, ns.task_id, title=ns.title,
                                             note=ns.note, due_date=ns.due,
                                             reminder_datetime=ns.reminder)
            print(f"Task updated: {result.get('title')}")

        elif cmd == "todo-task-complete":
            if len(args) < 2:
                print("Usage: m365 todo-task-complete <list-id> <task-id>")
                sys.exit(1)
            result = client.todo_complete_task(args[0], args[1])
            print(f"Task marked as completed: {result.get('title', args[1])}")

        elif cmd == "todo-task-delete":
            if len(args) < 2:
                print("Usage: m365 todo-task-delete <list-id> <task-id>")
                sys.exit(1)
            print(client.todo_delete_task(args[0], args[1]))

        elif cmd == "todo-step-add":
            if len(args) < 3:
                print("Usage: m365 todo-step-add <list-id> <task-id> <title>")
                sys.exit(1)
            result = client.todo_add_step(args[0], args[1], " ".join(args[2:]))
            print(f"Step added: {result.get('displayName')} (id: {result.get('id')})")

        elif cmd == "todo-step-complete":
            if len(args) < 3:
                print("Usage: m365 todo-step-complete <list-id> <task-id> <step-id>")
                sys.exit(1)
            client.todo_complete_step(args[0], args[1], args[2])
            print(f"Step {args[2]} marked as completed")

        elif cmd == "todo-task-assign":
            if len(args) < 3:
                print("Usage: m365 todo-task-assign <list-id> <task-id> <email>")
                sys.exit(1)
            result = client.todo_assign_task(args[0], args[1], args[2])
            print(f"Task assigned to {args[2]}: {result.get('title', args[1])}")

        elif cmd == "todo-task-move":
            if len(args) < 3:
                print("Usage: m365 todo-task-move <source-list-id> <task-id> <target-list-id>")
                sys.exit(1)
            result = client.todo_move_task(args[0], args[1], args[2])
            print(f"Task moved: {result.get('title')} → list {args[2]}")

        else:
            print(f"Unknown command: {cmd}", file=sys.stderr)
            print(USAGE)
            sys.exit(1)

    except PermissionError as exc:
        print(f"Permission Error: {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
PYEOF

# 9. Create SKILL.md for OpenClaw agent auto-discovery
cat > "${INSTALL_DIR}/SKILL.md" << 'EOF'
# SKILL: m365-graph

OpenClaw skill for unattended Microsoft 365 access via Graph API.

## Executable
m365

## Description
Enables OpenClaw agents to send/read mail, manage calendar events, manage contacts,
access OneDrive files, create OneNote pages, edit Excel/Word/PowerPoint documents,
and manage Microsoft To Do tasks. Uses client-credentials (daemon) flow for most
operations; OneNote requires delegated flow.

## Configuration
Credentials stored in: ~/.openclaw/skills/m365-graph/.env

Required environment variables:
- TENANT_ID
- CLIENT_ID
- CLIENT_SECRET
- TOKEN_CACHE_PATH
- M365_USER_EMAIL

## Required Microsoft Graph Permissions

**Application permissions** (unattended/daemon access):
`Mail.ReadWrite.All`, `Mail.Send`, `Calendars.ReadWrite.All`,
`Files.ReadWrite.All`, `Contacts.ReadWrite`, `Tasks.ReadWrite`

**Delegated permissions** (required for OneNote):
`Notes.ReadWrite.All`, `offline_access`

## Commands

| Command | Description | Example |
|---------|-------------|---------|
| mail-list [N] | List N recent inbox messages (default 20) | m365 mail-list 10 |
| send-mail \<to\> \<subject\> \<body\> | Send email (basic) | m365 send-mail boss@company.com "Update" "Project on track" |
| send-mail \<to\> \<subject\> \<body\> [options] | Send email with CC/BCC/attachments/priority | m365 send-mail boss@company.com "Update" "See attached" --cc cfo@company.com --importance High --attach /path/report.pdf |
| mail-read-subject \<subject\> | Read mail by subject search | m365 mail-read-subject "Invoice" |
| mail-read \<message-id\> | Read a specific message by ID | m365 mail-read AAMkAGI... |
| mail-reply \<message-id\> \<body\> | Reply to sender | m365 mail-reply AAMkAGI... "Got it, thanks!" |
| mail-reply-all \<message-id\> \<body\> | Reply to all | m365 mail-reply-all AAMkAGI... "All noted" |
| mail-forward \<message-id\> \<to\> [body] | Forward message | m365 mail-forward AAMkAGI... colleague@company.com "FYI" |
| mail-delete \<message-id\> | Delete a message | m365 mail-delete AAMkAGI... |
| mail-move \<message-id\> \<folder\> | Move message to folder | m365 mail-move AAMkAGI... Archive |
| calendar-list [days] | List upcoming events | m365 calendar-list 14 |
| calendar-create \<subj\> \<start\> \<end\> [body] | Create calendar event | m365 calendar-create "Meeting" 2026-04-15T10:00:00 2026-04-15T11:00:00 "Q1 review" |
| calendar-create ... --required email | Add required attendee | m365 calendar-create "Sync" 2026-04-15T10:00:00 2026-04-15T11:00:00 --required alice@co.com |
| calendar-create ... --optional email | Add optional attendee | m365 calendar-create "Sync" ... --optional bob@co.com |
| calendar-create ... --private | Mark as private | m365 calendar-create "1:1" ... --private |
| calendar-create ... --location addr | Set location | m365 calendar-create "Meeting" ... --location "Room 101" |
| calendar-create ... --reminder-minutes N | Set reminder | m365 calendar-create "Call" ... --reminder-minutes 30 |
| calendar-create ... --attach filepath | Add attachment | m365 calendar-create "Review" ... --attach /path/agenda.pdf |
| contacts-list [N] | List contacts (first,last,biz-mail,priv-mail,biz-phone,mobile) | m365 contacts-list |
| contacts-get \<id\> | Get all fields of a contact | m365 contacts-get AAMkAGI... |
| contacts-create \<first\> \<last\> [options] | Create contact | m365 contacts-create Jane Doe --email-business j@work.com --phone-mobile +1555 |
| contacts-photo-update \<id\> \<path\> | Set contact photo | m365 contacts-photo-update AAMkAGI... /path/photo.jpg |
| contacts-photo-delete \<id\> | Delete contact photo | m365 contacts-photo-delete AAMkAGI... |
| onedrive-list [folder] | List OneDrive folder | m365 onedrive-list /Documents |
| upload \<local\> \<remote\> | Upload to OneDrive | m365 upload ./report.pdf /Finance/report.pdf |
| download \<remote\> \<local\> | Download from OneDrive | m365 download /Finance/report.pdf ./local.pdf |
| onenote-create \<nb\> \<sec\> \<title\> \<html\> | Create OneNote page | m365 onenote-create WorkNotes April "Sync" "\<h1\>Notes\</h1\>" |
| excel-update \<path\> \<sheet\> \<range\> [val ...] | Update Excel cells | m365 excel-update Report.xlsx Sheet1 A1:B2 Jan Feb 100 200 |
| word-update \<path\> key=val ... | Replace placeholders in Word | m365 word-update Offer.docx "{{Name}}=Alice" |
| ppt-update \<path\> \<slide\> key=val ... | Replace text in PowerPoint | m365 ppt-update Deck.pptx 0 "{{Title}}=2026 Strategy" |
| todo-lists | List all To Do task lists | m365 todo-lists |
| todo-list-create \<name\> | Create task list | m365 todo-list-create "Work Tasks" |
| todo-list-rename \<id\> \<name\> | Rename task list | m365 todo-list-rename AAMk... "New Name" |
| todo-list-delete \<id\> | Delete task list | m365 todo-list-delete AAMk... |
| todo-tasks \<list-id\> | List tasks in a list | m365 todo-tasks AAMk... |
| todo-tasks-all | List all tasks | m365 todo-tasks-all |
| todo-task-create \<list-id\> \<title\> [opts] | Create task | m365 todo-task-create AAMk... "Buy groceries" --due 2026-04-15 |
| todo-task-update \<list-id\> \<task-id\> [opts] | Update task | m365 todo-task-update AAMk... AAMk2... --title "Buy organic groceries" |
| todo-task-complete \<list-id\> \<task-id\> | Complete task | m365 todo-task-complete AAMk... AAMk2... |
| todo-task-delete \<list-id\> \<task-id\> | Delete task | m365 todo-task-delete AAMk... AAMk2... |
| todo-step-add \<list-id\> \<task-id\> \<title\> | Add step to task | m365 todo-step-add AAMk... AAMk2... "Buy milk" |
| todo-step-complete \<list-id\> \<task-id\> \<step-id\> | Complete step | m365 todo-step-complete AAMk... AAMk2... AAMk3... |
| todo-task-assign \<list-id\> \<task-id\> \<email\> | Assign task | m365 todo-task-assign AAMk... AAMk2... alice@co.com |
| todo-task-move \<src-list\> \<task-id\> \<dst-list\> | Move task | m365 todo-task-move AAMk... AAMk2... AAMk4... |

## Notes
- Mail access requires Exchange Online Application Access Policy in addition to Mail.ReadWrite.All
- Tasks.ReadWrite permission required for ToDo features (add to app registration)
- OneNote requires delegated Notes.ReadWrite.All permission
- Excel update uses Graph API directly; file must not be open in Office
- Word and PowerPoint use download → edit locally → re-upload workflow

EOF
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
echo "--- Verifying mail access (requires Mail.ReadWrite.All + Exchange Online policy) ---"
mail_output=$("${VENV_DIR}/bin/python" -m m365_openclaw mail-list 2>&1)
mail_exit=$?
if [ "${mail_exit}" -eq 0 ]; then
    echo "${mail_output}"
    echo "Mail access: OK"
else
    echo "${mail_output}" >&2
    echo ""
    echo "WARNING: Mail access failed (see error above)."
    echo "  This is usually caused by one of:"
    echo "    1. 'Mail.ReadWrite.All' application permission not admin-consented."
    echo "       Go to: Entra ID → App registrations → <your app> → API permissions"
    echo "    2. Exchange Online Application Access Policy not configured."
    echo "       Ask your Exchange admin to run in Exchange Online PowerShell:"
    echo "         New-ApplicationAccessPolicy \\"
    echo "           -AppId ${CLIENT_ID} \\"
    echo "           -PolicyScopeGroupId ${M365_USER} \\"
    echo "           -AccessRight RestrictAccess \\"
    echo "           -Description 'OpenClaw M365 mail access'"
    echo "       Reference: https://aka.ms/graph-app-access-policy"
    echo "  Calendar and contacts are unaffected and will continue to work."
fi

echo ""
echo "=== Installation complete! ==="
echo ""
echo "Available commands:"
echo "  m365 mail-list [N]             – list recent inbox messages"
echo "  m365 send-mail <to> <sub> <body> [opts] – send email (--cc, --bcc, --attach, --importance)"
echo "  m365 mail-read-subject <subj>  – find messages by subject"
echo "  m365 mail-read <id>            – read a message"
echo "  m365 mail-reply <id> <body>    – reply to sender"
echo "  m365 mail-reply-all <id> <body> – reply to all"
echo "  m365 mail-forward <id> <to>    – forward a message"
echo "  m365 mail-delete <id>          – delete a message"
echo "  m365 mail-move <id> <folder>   – move to folder"
echo "  m365 calendar-list [days]      – list upcoming calendar events"
echo "  m365 calendar-create <subj> <start> <end> [body] [opts] – create event"
echo "  m365 contacts-list [N]         – list contacts"
echo "  m365 contacts-get <id>         – get all contact fields"
echo "  m365 contacts-create <first> <last> [opts] – add a contact"
echo "  m365 contacts-photo-update <id> <path> – update contact photo"
echo "  m365 contacts-photo-delete <id> – delete contact photo"
echo "  m365 onedrive-list [folder]    – list OneDrive folder"
echo "  m365 upload <local> <remote>   – upload file to OneDrive"
echo "  m365 download <remote> <local> – download file from OneDrive"
echo "  m365 onenote-create ...        – create a OneNote page"
echo "  m365 excel-update ...          – update Excel cells"
echo "  m365 word-update ...           – replace text in Word document"
echo "  m365 ppt-update ...            – replace text in PowerPoint slide"
echo "  m365 todo-lists                – list all task lists"
echo "  m365 todo-list-create <name>   – create a task list"
echo "  m365 todo-list-rename <id> <name> – rename a task list"
echo "  m365 todo-list-delete <id>     – delete a task list"
echo "  m365 todo-tasks <list-id>      – list tasks in a list"
echo "  m365 todo-tasks-all            – list all tasks"
echo "  m365 todo-task-create <list-id> <title> [opts] – create task"
echo "  m365 todo-task-update <list-id> <task-id> [opts] – update task"
echo "  m365 todo-task-complete <list-id> <task-id> – complete a task"
echo "  m365 todo-task-delete <list-id> <task-id>  – delete a task"
echo "  m365 todo-step-add <list-id> <task-id> <title> – add step"
echo "  m365 todo-step-complete <list-id> <task-id> <step-id> – complete step"
echo "  m365 todo-task-assign <list-id> <task-id> <email> – assign task"
echo "  m365 todo-task-move <src-list> <task-id> <dst-list> – move task"
echo ""
echo "If authentication fails or no calendar appears:"
echo "  rm -rf ${OPENCLAW_DIR}/credentials/m365_token_cache.bin"
echo "  cd ${INSTALL_DIR}"
echo "  ${VENV_DIR}/bin/python -m m365_openclaw calendar-list"
echo ""
echo "Done!"
