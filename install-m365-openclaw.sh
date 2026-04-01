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
    version='0.2.0',
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
echo "                Files.ReadWrite.All, Contacts.ReadWrite, Tasks.ReadWrite.All"
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
core.py – Microsoft 365 client for OpenClaw agents (v0.2.0).

Supports: Mail, Calendar, Contacts, OneDrive, OneNote, Excel, Word, PowerPoint,
          Microsoft ToDo tasks.
Uses client-credentials (daemon/application) flow via the O365 library + MSAL.

Important: When using application permissions (client_credentials flow), all
Graph API calls must target a specific user.  Set M365_USER_EMAIL in .env so
that main_resource resolves /me/ to /users/<email>/ automatically.
"""

import base64
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

    def _graph_headers(self, content_type='application/json'):
        headers = {'Authorization': f'Bearer {self._access_token()}'}
        if content_type:
            headers['Content-Type'] = content_type
        return headers

    def _graph_get(self, url, params=None):
        """HTTP GET against the Microsoft Graph API."""
        import requests as req
        resp = req.get(url, headers=self._graph_headers(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_post(self, url, data=None, content_type='application/json', raw_data=None):
        """HTTP POST to the Microsoft Graph API."""
        import requests as req
        headers = self._graph_headers(content_type=content_type)
        if raw_data is not None:
            resp = req.post(url, headers=headers, data=raw_data, timeout=30)
        else:
            resp = req.post(url, headers=headers, json=data, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_patch(self, url, data):
        """HTTP PATCH to the Microsoft Graph API with JSON body."""
        import requests as req
        resp = req.patch(url, headers=self._graph_headers(), json=data, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_delete(self, url):
        """HTTP DELETE to the Microsoft Graph API."""
        import requests as req
        resp = req.delete(url, headers=self._graph_headers(content_type=None), timeout=30)
        resp.raise_for_status()
        return {}

    def _graph_put(self, url, data, content_type='application/octet-stream'):
        """HTTP PUT to the Microsoft Graph API (used for binary uploads)."""
        import requests as req
        resp = req.put(url, headers=self._graph_headers(content_type=content_type),
                       data=data, timeout=60)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _base_url(self):
        """Return the per-user Graph API base URL."""
        return f"https://graph.microsoft.com/v1.0/users/{self.user}"

    # ── Mail – error helper ───────────────────────────────────────────────────

    @staticmethod
    def _raise_if_mail_403(exc):
        """Re-raise with actionable guidance when a 403 occurs on a mail endpoint."""
        try:
            import requests
            if isinstance(exc, requests.exceptions.HTTPError):
                resp = getattr(exc, 'response', None)
                if resp is not None and resp.status_code == 403:
                    url = getattr(resp, 'url', '') or ''
                    if 'mailFolders' in url or '/messages' in url or 'sendMail' in url:
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
        except ImportError:
            pass

    # ── Mail ──────────────────────────────────────────────────────────────────

    def send_mail(self, to_address, subject, body, cc=None, bcc=None,
                  sensitivity='Normal', importance='Normal', attachments=None,
                  request_delivery_receipt=False, request_read_receipt=False):
        """
        Send an email via the Graph API with full feature support.

        Args:
            to_address: Recipient address string or list of strings.
            subject:    Email subject.
            body:       Email body (plain text or HTML).
            cc:         CC address string or list of strings (optional).
            bcc:        BCC address string or list of strings (optional).
            sensitivity: Normal | Personal | Private | Confidential (default: Normal).
            importance:  High | Normal | Low (default: Normal).
            attachments: Local file path string or list of paths (optional).
            request_delivery_receipt: Request delivery receipt (default: False).
            request_read_receipt:     Request read receipt (default: False).
        """
        def _recipients(addrs):
            if not addrs:
                return []
            if isinstance(addrs, str):
                addrs = [addrs]
            return [{'emailAddress': {'address': a.strip()}} for a in addrs if a.strip()]

        message = {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': body},
            'toRecipients': _recipients(to_address),
            'importance': importance,
            'sensitivity': sensitivity,
            'isDeliveryReceiptRequested': request_delivery_receipt,
            'isReadReceiptRequested': request_read_receipt,
        }
        if cc:
            message['ccRecipients'] = _recipients(cc)
        if bcc:
            message['bccRecipients'] = _recipients(bcc)

        # Add file attachments encoded as base64
        if attachments:
            att_list = attachments if isinstance(attachments, list) else [attachments]
            graph_atts = []
            for att_path in att_list:
                with open(att_path, 'rb') as fh:
                    content_b64 = base64.b64encode(fh.read()).decode()
                graph_atts.append({
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    'name': Path(att_path).name,
                    'contentBytes': content_b64,
                })
            message['attachments'] = graph_atts

        url = f"{self._base_url()}/sendMail"
        try:
            self._graph_post(url, {'message': message, 'saveToSentItems': True})
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise

        recipients = to_address if isinstance(to_address, list) else [to_address]
        return f"Email sent to {', '.join(recipients)}"

    def list_mail(self, limit=20):
        """Return recent messages from the inbox, including message IDs for follow-up."""
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

    def search_mail(self, subject_query, limit=10):
        """
        Search inbox messages by subject keyword.
        Returns a list of matching messages including their IDs, body content,
        and sender details – ready for follow-up actions (reply, forward, etc.).
        """
        url = (
            f"{self._base_url()}/mailFolders/inbox/messages"
            f"?$top={limit}"
            f"&$select=id,subject,from,receivedDateTime,isRead,body"
        )
        # OData 'contains' filter on subject
        url += f"&$filter=contains(subject,'{subject_query}')"
        try:
            data = self._graph_get(url)
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        result = []
        for msg in data.get('value', []):
            result.append({
                'id': msg.get('id', ''),
                'subject': msg.get('subject', '(no subject)'),
                'from_address': msg.get('from', {}).get('emailAddress', {}).get('address', ''),
                'from_name': msg.get('from', {}).get('emailAddress', {}).get('name', ''),
                'date': msg.get('receivedDateTime', ''),
                'is_read': msg.get('isRead', False),
                'body': msg.get('body', {}).get('content', ''),
            })
        return result

    def get_mail_headers(self, message_id):
        """Return headers and metadata for a specific message by ID."""
        url = (
            f"{self._base_url()}/messages/{message_id}"
            f"?$select=id,subject,from,toRecipients,ccRecipients,"
            f"receivedDateTime,sentDateTime,importance,sensitivity,"
            f"isRead,hasAttachments"
        )
        try:
            msg = self._graph_get(url)
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return {
            'id': msg.get('id', ''),
            'subject': msg.get('subject', '(no subject)'),
            'from': msg.get('from', {}).get('emailAddress', {}),
            'to': [r.get('emailAddress', {}) for r in msg.get('toRecipients', [])],
            'cc': [r.get('emailAddress', {}) for r in msg.get('ccRecipients', [])],
            'date_received': msg.get('receivedDateTime', ''),
            'date_sent': msg.get('sentDateTime', ''),
            'importance': msg.get('importance', ''),
            'sensitivity': msg.get('sensitivity', ''),
            'is_read': msg.get('isRead', False),
            'has_attachments': msg.get('hasAttachments', False),
        }

    def reply_mail(self, message_id, body):
        """Reply to the sender of a message (reply to sender only)."""
        url = f"{self._base_url()}/messages/{message_id}/reply"
        try:
            self._graph_post(url, {'message': {}, 'comment': body})
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Reply sent for message {message_id}"

    def reply_all_mail(self, message_id, body):
        """Reply to all recipients of a message."""
        url = f"{self._base_url()}/messages/{message_id}/replyAll"
        try:
            self._graph_post(url, {'message': {}, 'comment': body})
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Reply-All sent for message {message_id}"

    def forward_mail(self, message_id, to_address, body=""):
        """Forward a message to one or more new recipients."""
        if isinstance(to_address, str):
            to_address = [to_address]
        url = f"{self._base_url()}/messages/{message_id}/forward"
        try:
            self._graph_post(url, {
                'comment': body,
                'toRecipients': [
                    {'emailAddress': {'address': a.strip()}} for a in to_address
                ],
            })
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Message forwarded to {', '.join(to_address)}"

    def delete_mail(self, message_id):
        """Permanently delete a message by ID."""
        url = f"{self._base_url()}/messages/{message_id}"
        try:
            self._graph_delete(url)
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Message {message_id} deleted"

    def move_mail(self, message_id, destination_folder):
        """
        Move a message to another mail folder.
        Use well-known names: inbox, sent, drafts, deleted, archive, junk.
        Or provide a folder ID directly.
        """
        well_known = {
            'inbox': 'inbox',
            'sent': 'sentitems',
            'drafts': 'drafts',
            'deleted': 'deleteditems',
            'archive': 'archive',
            'junk': 'junkemail',
        }
        folder_id = well_known.get(destination_folder.lower(), destination_folder)
        url = f"{self._base_url()}/messages/{message_id}/move"
        try:
            result = self._graph_post(url, {'destinationId': folder_id})
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        new_id = result.get('id', message_id)
        return f"Message moved to '{destination_folder}' (new ID: {new_id})"

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
                               is_private=False, reminder_minutes=None, attachment=None):
        """
        Create a calendar event via Graph API with extended options.

        Args:
            subject:             Event title.
            start_iso:           Start time in ISO 8601 format (e.g. 2026-04-15T10:00:00).
            end_iso:             End time in ISO 8601 format.
            body:                Event description/body text (optional).
            location:            Location or address (optional).
            required_attendees:  Email string or list of emails for required attendees.
            optional_attendees:  Email string or list of emails for optional attendees.
            is_private:          Mark event as private (default: False).
            reminder_minutes:    Minutes before event to trigger a reminder (optional).
            attachment:          Local file path to attach to the event (optional).
        """
        event_data = {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': body},
            'start': {'dateTime': start_iso, 'timeZone': 'UTC'},
            'end': {'dateTime': end_iso, 'timeZone': 'UTC'},
            'sensitivity': 'private' if is_private else 'normal',
        }
        if location:
            event_data['location'] = {'displayName': location}

        # Build attendees list
        attendees = []
        if required_attendees:
            if isinstance(required_attendees, str):
                required_attendees = [required_attendees]
            for email in required_attendees:
                attendees.append({'emailAddress': {'address': email.strip()}, 'type': 'required'})
        if optional_attendees:
            if isinstance(optional_attendees, str):
                optional_attendees = [optional_attendees]
            for email in optional_attendees:
                attendees.append({'emailAddress': {'address': email.strip()}, 'type': 'optional'})
        if attendees:
            event_data['attendees'] = attendees

        # Reminder
        if reminder_minutes is not None:
            event_data['isReminderOn'] = True
            event_data['reminderMinutesBeforeStart'] = int(reminder_minutes)

        # File attachment (base64)
        if attachment:
            with open(attachment, 'rb') as fh:
                content_b64 = base64.b64encode(fh.read()).decode()
            event_data['attachments'] = [{
                '@odata.type': '#microsoft.graph.fileAttachment',
                'name': Path(attachment).name,
                'contentBytes': content_b64,
            }]

        url = f"{self._base_url()}/calendar/events"
        result = self._graph_post(url, event_data)
        event_id = result.get('id', '')
        return f"Event '{subject}' created (ID: {event_id})"

    # ── Contacts ──────────────────────────────────────────────────────────────

    def list_contacts(self, limit=100):
        """
        List contacts. Returns columns: ID, first name, last name, work email,
        personal email, work phone, home phone, mobile phone.
        """
        url = (
            f"{self._base_url()}/contacts"
            f"?$top={limit}"
            f"&$select=id,givenName,surname,displayName,emailAddresses,"
            f"businessPhones,homePhones,mobilePhone"
        )
        data = self._graph_get(url)
        result = []
        for c in data.get('value', []):
            emails = c.get('emailAddresses', [])
            # First email is treated as work, second as home if no type label
            work_email = ''
            home_email = ''
            for e in emails:
                name_lc = e.get('name', '').lower()
                addr = e.get('address', '')
                if name_lc in ('work', 'geschäft', 'business') and not work_email:
                    work_email = addr
                elif name_lc in ('home', 'personal', 'privat') and not home_email:
                    home_email = addr
                elif not work_email:
                    work_email = addr
                elif not home_email:
                    home_email = addr
            result.append({
                'id': c.get('id', ''),
                'first_name': c.get('givenName', ''),
                'last_name': c.get('surname', ''),
                'display_name': c.get('displayName', ''),
                'work_email': work_email,
                'home_email': home_email,
                'work_phone': (c.get('businessPhones') or [''])[0],
                'home_phone': (c.get('homePhones') or [''])[0],
                'mobile_phone': c.get('mobilePhone', ''),
            })
        return result

    def get_contact(self, name):
        """
        Get all available fields for the first contact whose display name contains *name*.
        Returns None if no match is found.
        """
        url = (
            f"{self._base_url()}/contacts"
            f"?$top=5"
            f"&$filter=contains(displayName,'{name}')"
        )
        data = self._graph_get(url)
        contacts = data.get('value', [])
        if not contacts:
            return None
        c = contacts[0]
        return {
            'id': c.get('id', ''),
            'display_name': c.get('displayName', ''),
            'first_name': c.get('givenName', ''),
            'last_name': c.get('surname', ''),
            'company': c.get('companyName', ''),
            'job_title': c.get('jobTitle', ''),
            'emails': c.get('emailAddresses', []),
            'business_phones': c.get('businessPhones', []),
            'home_phones': c.get('homePhones', []),
            'mobile_phone': c.get('mobilePhone', ''),
            'business_address': c.get('businessAddress', {}),
            'home_address': c.get('homeAddress', {}),
            'birthday': c.get('birthday', ''),
            'anniversary': c.get('anniversary', ''),
            'spouse_name': c.get('spouseName', ''),
            'websites': c.get('websites', []),
            'personal_notes': c.get('personalNotes', ''),
        }

    def create_contact(self, given_name, surname, email=None, phone=None,
                        work_email=None, home_email=None,
                        work_phone=None, home_phone=None, mobile_phone=None,
                        work_street=None, work_city=None, work_state=None,
                        work_zip=None, work_country=None,
                        home_street=None, home_city=None, home_state=None,
                        home_zip=None, home_country=None,
                        birthday=None, anniversary=None,
                        website=None, work_website=None,
                        spouse=None, notes=None):
        """
        Create a new contact with extended optional fields.

        Backward compatible: positional 'email' and 'phone' still work as work
        email and work phone respectively.  All additional fields are optional flags.

        Note: Fields such as hobbies, zodiac sign, and children count are not
        supported by the Graph API contacts schema.  Include them in the *notes*
        parameter to preserve the information.
        """
        # Build email address list
        email_list = []
        resolved_work_email = work_email or email
        if resolved_work_email:
            email_list.append({'name': 'Work', 'address': resolved_work_email.strip()})
        if home_email:
            email_list.append({'name': 'Home', 'address': home_email.strip()})

        contact_data = {
            'givenName': given_name,
            'surname': surname,
            'emailAddresses': email_list,
        }

        # Business phones
        resolved_work_phone = work_phone or phone
        if resolved_work_phone:
            contact_data['businessPhones'] = [resolved_work_phone.strip()]
        if home_phone:
            contact_data['homePhones'] = [home_phone.strip()]
        if mobile_phone:
            contact_data['mobilePhone'] = mobile_phone.strip()

        # Addresses
        if any([work_street, work_city, work_state, work_zip, work_country]):
            contact_data['businessAddress'] = {
                'street': work_street or '',
                'city': work_city or '',
                'state': work_state or '',
                'postalCode': work_zip or '',
                'countryOrRegion': work_country or '',
            }
        if any([home_street, home_city, home_state, home_zip, home_country]):
            contact_data['homeAddress'] = {
                'street': home_street or '',
                'city': home_city or '',
                'state': home_state or '',
                'postalCode': home_zip or '',
                'countryOrRegion': home_country or '',
            }

        # Personal details
        if birthday:
            contact_data['birthday'] = birthday        # ISO 8601, e.g. 1990-05-15T00:00:00Z
        if anniversary:
            contact_data['anniversary'] = anniversary
        if spouse:
            contact_data['spouseName'] = spouse
        if notes:
            contact_data['personalNotes'] = notes

        # Websites
        websites = []
        if website:
            websites.append({'type': 'home', 'address': website})
        if work_website:
            websites.append({'type': 'work', 'address': work_website})
        if websites:
            contact_data['websites'] = websites

        url = f"{self._base_url()}/contacts"
        result = self._graph_post(url, contact_data)
        contact_id = result.get('id', '')
        return f"Contact '{given_name} {surname}' created (ID: {contact_id})"

    def set_contact_photo(self, contact_id, photo_path):
        """Upload a photo for a contact (JPEG recommended, max 4 MB)."""
        with open(photo_path, 'rb') as fh:
            photo_bytes = fh.read()
        url = f"{self._base_url()}/contacts/{contact_id}/photo/$value"
        self._graph_put(url, photo_bytes, content_type='image/jpeg')
        return f"Photo set for contact {contact_id}"

    def delete_contact_photo(self, contact_id):
        """Delete the profile photo for a contact."""
        url = f"{self._base_url()}/contacts/{contact_id}/photo/$value"
        self._graph_delete(url)
        return f"Photo deleted for contact {contact_id}"

    def get_contact_photo(self, contact_id, save_path):
        """Download the profile photo of a contact to a local file."""
        import requests as req
        url = f"{self._base_url()}/contacts/{contact_id}/photo/$value"
        resp = req.get(url, headers=self._graph_headers(content_type=None), timeout=30)
        resp.raise_for_status()
        with open(save_path, 'wb') as fh:
            fh.write(resp.content)
        return f"Photo saved to {save_path}"

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

    # ── Microsoft ToDo / Tasks ────────────────────────────────────────────────

    def todo_list_task_lists(self):
        """List all Microsoft ToDo task lists. Returns list with id and name."""
        url = f"{self._base_url()}/todo/lists"
        data = self._graph_get(url)
        return [
            {
                'id': lst.get('id', ''),
                'name': lst.get('displayName', ''),
                'is_owner': lst.get('isOwner', True),
                'is_shared': lst.get('isShared', False),
            }
            for lst in data.get('value', [])
        ]

    def todo_create_task_list(self, name):
        """Create a new Microsoft ToDo task list with the given name."""
        url = f"{self._base_url()}/todo/lists"
        result = self._graph_post(url, {'displayName': name})
        return f"Task list '{name}' created (ID: {result.get('id', '')})"

    def todo_rename_task_list(self, list_id, new_name):
        """Rename an existing task list."""
        url = f"{self._base_url()}/todo/lists/{list_id}"
        self._graph_patch(url, {'displayName': new_name})
        return f"Task list {list_id} renamed to '{new_name}'"

    def todo_delete_task_list(self, list_id):
        """Delete a task list and all its tasks permanently."""
        url = f"{self._base_url()}/todo/lists/{list_id}"
        self._graph_delete(url)
        return f"Task list {list_id} deleted"

    def todo_list_tasks(self, list_id, due_after=None):
        """
        List tasks in a specific task list.
        Args:
            list_id:   ID of the task list.
            due_after: Optional ISO 8601 date (e.g. 2026-01-01) to show only tasks
                       due on or after that date (client-side filter).
        """
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks?$top=100&$expand=checklistItems"
        data = self._graph_get(url)
        tasks = self._format_tasks(data.get('value', []))
        if due_after:
            tasks = [t for t in tasks if t.get('due', '') >= due_after]
        return tasks

    def todo_get_all_tasks(self, due_after=None):
        """List all tasks across every task list, optionally filtered by due date."""
        lists_data = self._graph_get(f"{self._base_url()}/todo/lists")
        all_tasks = []
        for lst in lists_data.get('value', []):
            list_id = lst['id']
            list_name = lst.get('displayName', '')
            tasks = self.todo_list_tasks(list_id, due_after=due_after)
            for task in tasks:
                task['list_name'] = list_name
                task['list_id'] = list_id
            all_tasks.extend(tasks)
        return all_tasks

    @staticmethod
    def _format_tasks(tasks):
        """Normalise raw Graph API task objects into a clean dict structure."""
        result = []
        for t in tasks:
            due = t.get('dueDateTime') or {}
            reminder = t.get('reminderDateTime') or {}
            result.append({
                'id': t.get('id', ''),
                'title': t.get('title', '(no title)'),
                'status': t.get('status', ''),
                'importance': t.get('importance', ''),
                'is_done': t.get('status', '') == 'completed',
                'due': due.get('dateTime', '') if due else '',
                'reminder': reminder.get('dateTime', '') if reminder else '',
                'note': (t.get('body') or {}).get('content', ''),
                'steps': [
                    {
                        'id': s.get('id', ''),
                        'title': s.get('displayName', ''),
                        'is_done': s.get('isChecked', False),
                    }
                    for s in (t.get('checklistItems') or [])
                ],
            })
        return result

    def todo_create_task(self, list_id, title, note=None, due_date=None,
                          reminder_datetime=None):
        """
        Create a new task in a task list.
        Args:
            list_id:           ID of the target task list.
            title:             Task title.
            note:              Optional task note/description.
            due_date:          Optional due date (ISO 8601, e.g. 2026-04-15T00:00:00).
            reminder_datetime: Optional reminder datetime (ISO 8601).
        """
        task_data = {'title': title}
        if note:
            task_data['body'] = {'content': note, 'contentType': 'text'}
        if due_date:
            task_data['dueDateTime'] = {'dateTime': due_date, 'timeZone': 'UTC'}
        if reminder_datetime:
            task_data['reminderDateTime'] = {'dateTime': reminder_datetime, 'timeZone': 'UTC'}
            task_data['isReminderOn'] = True
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks"
        result = self._graph_post(url, task_data)
        return f"Task '{title}' created (ID: {result.get('id', '')})"

    def todo_update_task(self, list_id, task_id, title=None, note=None,
                          due_date=None, reminder_datetime=None):
        """Update one or more fields of an existing task."""
        task_data = {}
        if title:
            task_data['title'] = title
        if note is not None:
            task_data['body'] = {'content': note, 'contentType': 'text'}
        if due_date:
            task_data['dueDateTime'] = {'dateTime': due_date, 'timeZone': 'UTC'}
        if reminder_datetime:
            task_data['reminderDateTime'] = {'dateTime': reminder_datetime, 'timeZone': 'UTC'}
            task_data['isReminderOn'] = True
        if not task_data:
            return "No fields to update – specify at least one option."
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks/{task_id}"
        self._graph_patch(url, task_data)
        return f"Task {task_id} updated"

    def todo_complete_task(self, list_id, task_id):
        """Mark a task as completed."""
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks/{task_id}"
        self._graph_patch(url, {'status': 'completed'})
        return f"Task {task_id} marked as completed"

    def todo_add_step(self, list_id, task_id, step_title):
        """Add a checklist step (subtask) to a task."""
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks/{task_id}/checklistItems"
        result = self._graph_post(url, {'displayName': step_title, 'isChecked': False})
        return f"Step '{step_title}' added (ID: {result.get('id', '')})"

    def todo_complete_step(self, list_id, task_id, step_id):
        """Mark an individual checklist step as completed."""
        url = (
            f"{self._base_url()}/todo/lists/{list_id}"
            f"/tasks/{task_id}/checklistItems/{step_id}"
        )
        self._graph_patch(url, {'isChecked': True})
        return f"Step {step_id} marked as completed"

    def todo_move_task(self, from_list_id, task_id, to_list_id):
        """
        Move a task from one list to another.
        The Graph API has no native 'move' endpoint; this copies the task
        (including checklist steps and status) then deletes the original.
        """
        # Fetch original task with its checklist items
        src_url = (
            f"{self._base_url()}/todo/lists/{from_list_id}"
            f"/tasks/{task_id}?$expand=checklistItems"
        )
        task = self._graph_get(src_url)

        # Build new task payload
        task_data = {'title': task.get('title', '')}
        if task.get('body'):
            task_data['body'] = task['body']
        if task.get('dueDateTime'):
            task_data['dueDateTime'] = task['dueDateTime']
        if task.get('reminderDateTime'):
            task_data['reminderDateTime'] = task['reminderDateTime']
            task_data['isReminderOn'] = task.get('isReminderOn', False)
        if task.get('status') == 'completed':
            task_data['status'] = 'completed'

        dst_url = f"{self._base_url()}/todo/lists/{to_list_id}/tasks"
        new_task = self._graph_post(dst_url, task_data)
        new_task_id = new_task.get('id', '')

        # Copy checklist steps to the new task
        for item in (task.get('checklistItems') or []):
            step_url = (
                f"{self._base_url()}/todo/lists/{to_list_id}"
                f"/tasks/{new_task_id}/checklistItems"
            )
            self._graph_post(step_url, {
                'displayName': item.get('displayName', ''),
                'isChecked': item.get('isChecked', False),
            })

        # Delete original task
        self._graph_delete(
            f"{self._base_url()}/todo/lists/{from_list_id}/tasks/{task_id}"
        )
        return (
            f"Task moved from list {from_list_id} to {to_list_id} "
            f"(new ID: {new_task_id})"
        )
PYEOF

cat > "${SRC_DIR}/m365_openclaw/__main__.py" << 'PYEOF'
"""
__main__.py – CLI entry point for the OpenClaw M365 skill (v0.2.0).

Invoked as:  python -m m365_openclaw <command> [arguments...]
Or via the  m365  wrapper script placed in ~/.local/bin.

All v0.1.0 commands remain 100% compatible.
"""

import sys

USAGE = """\
OpenClaw M365 CLI v0.2.0 – Microsoft 365 for agents

━━━ MAIL ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  mail-list [N]
      List N recent inbox messages with IDs (default: 20)

  send-mail <to> <subject> <body>
            [--cc addr,...] [--bcc addr,...] [--importance High|Normal|Low]
            [--sensitivity Normal|Personal|Private|Confidential]
            [--attach /path/to/file] [--delivery-receipt] [--read-receipt]
      Send an email. Multiple --attach flags allowed.

  mail-search <query>
      Search inbox by subject keyword; returns messages with IDs.

  mail-headers <message_id>
      Show headers/metadata of a specific message.

  mail-reply <message_id> <body>
      Reply to the sender of a message.

  mail-reply-all <message_id> <body>
      Reply to all recipients of a message.

  mail-forward <message_id> <to> [body]
      Forward a message to a new recipient.

  mail-delete <message_id>
      Permanently delete a message.

  mail-move <message_id> <folder>
      Move a message to another folder.
      Well-known folders: inbox, sent, drafts, deleted, archive, junk

━━━ CALENDAR ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  calendar-list [days]
      List upcoming events (default: 7 days)

  calendar-create <subject> <start_iso> <end_iso> [body]
                  [--attendees email,...] [--optional-attendees email,...]
                  [--location "Place"] [--private] [--reminder <minutes>]
                  [--attach /path/to/file]
      Create a calendar event (ISO 8601, e.g. 2026-04-15T10:00:00)

━━━ CONTACTS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  contacts-list [N]
      List N contacts as a table: First, Last, Work Mail, Home Mail,
      Work Phone, Home Phone, Mobile (default: 100)

  contacts-get <name>
      Show all fields for the first contact whose name contains <name>.

  contacts-create <first> <last> [email] [phone]
                  [--home-email e] [--home-phone p] [--mobile-phone p]
                  [--work-street s] [--work-city c] [--work-zip z] [--work-country c]
                  [--home-street s] [--home-city c] [--home-zip z] [--home-country c]
                  [--birthday YYYY-MM-DDT00:00:00Z] [--anniversary YYYY-MM-DDT00:00:00Z]
                  [--website URL] [--work-website URL] [--spouse name] [--notes text]
      Add a new contact. email and phone set work email/phone.

  contacts-photo-set <contact_id> <photo_path>
      Upload a profile photo (JPEG) for a contact.

  contacts-photo-delete <contact_id>
      Remove the profile photo for a contact.

  contacts-photo-get <contact_id> <save_path>
      Download the profile photo of a contact to a local file.

━━━ ONEDRIVE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  onedrive-list [folder]
      List files in a OneDrive folder (default: /)

  upload <local_path> <remote_path>
      Upload a local file to OneDrive

  download <remote_path> <local_path>
      Download a file from OneDrive

━━━ ONENOTE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  onenote-create <notebook> <section> <title> <html>
      Create a OneNote page (requires Notes.ReadWrite.All delegated permission)

━━━ EXCEL / WORD / POWERPOINT ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  excel-update <remote_path> <sheet> <range> [value ...]
      Update Excel cells, e.g.: excel-update report.xlsx Sheet1 A1:B2 Jan Feb 100 200

  word-update <remote_path> key=value [key=value ...]
      Replace placeholders in a Word document

  ppt-update <remote_path> <slide_number> key=value [key=value ...]
      Replace text in a PowerPoint slide (0-indexed)

━━━ MICROSOFT TODO ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  todo-list-lists
      List all ToDo task lists with their IDs.

  todo-create-list <name>
      Create a new task list.

  todo-rename-list <list_id> <new_name>
      Rename an existing task list.

  todo-delete-list <list_id>
      Delete a task list and all its tasks.

  todo-list-tasks <list_id> [--due-after YYYY-MM-DD]
      List tasks in a specific list (optionally filtered by due date).

  todo-all-tasks [--due-after YYYY-MM-DD]
      List all tasks across every task list.

  todo-create-task <list_id> <title>
                   [--note text] [--due YYYY-MM-DDT00:00:00]
                   [--reminder YYYY-MM-DDTHH:MM:SS]
      Create a new task in a list.

  todo-update-task <list_id> <task_id>
                   [--title text] [--note text] [--due YYYY-MM-DDT00:00:00]
                   [--reminder YYYY-MM-DDTHH:MM:SS]
      Update fields of an existing task.

  todo-complete-task <list_id> <task_id>
      Mark a task as completed.

  todo-add-step <list_id> <task_id> <step_title>
      Add a checklist step (subtask) to a task.

  todo-complete-step <list_id> <task_id> <step_id>
      Mark an individual step as completed.

  todo-move-task <from_list_id> <task_id> <to_list_id>
      Move a task to a different list (copy + delete).
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
        # ── Mail ─────────────────────────────────────────────────────────────

        if cmd == "mail-list":
            limit = int(args[0]) if args else 20
            msgs = client.list_mail(limit=limit)
            if not msgs:
                print("No messages found.")
            else:
                for m in msgs:
                    tag = "" if m['is_read'] else "[NEW] "
                    print(f"{tag}{m['date']} | ID: {m['id']} | {m['from']}: {m['subject']}")

        elif cmd == "send-mail":
            import argparse
            parser = argparse.ArgumentParser(prog='m365 send-mail')
            parser.add_argument('to', help='Recipient address(es), comma-separated')
            parser.add_argument('subject', help='Email subject')
            parser.add_argument('body', help='Email body (plain text or HTML)')
            parser.add_argument('--cc', default='', help='CC addresses, comma-separated')
            parser.add_argument('--bcc', default='', help='BCC addresses, comma-separated')
            parser.add_argument('--importance', choices=['High', 'Normal', 'Low'],
                                default='Normal')
            parser.add_argument('--sensitivity',
                                choices=['Normal', 'Personal', 'Private', 'Confidential'],
                                default='Normal')
            parser.add_argument('--attach', action='append', metavar='FILE', default=[],
                                help='Attach a local file (repeat for multiple files)')
            parser.add_argument('--delivery-receipt', action='store_true',
                                help='Request delivery receipt')
            parser.add_argument('--read-receipt', action='store_true',
                                help='Request read receipt')
            pargs = parser.parse_args(args)
            to_list = [a.strip() for a in pargs.to.split(',') if a.strip()]
            cc_list = [a.strip() for a in pargs.cc.split(',') if a.strip()]
            bcc_list = [a.strip() for a in pargs.bcc.split(',') if a.strip()]
            print(client.send_mail(
                to_list if len(to_list) > 1 else to_list[0],
                pargs.subject,
                pargs.body,
                cc=cc_list or None,
                bcc=bcc_list or None,
                importance=pargs.importance,
                sensitivity=pargs.sensitivity,
                attachments=pargs.attach or None,
                request_delivery_receipt=pargs.delivery_receipt,
                request_read_receipt=pargs.read_receipt,
            ))

        elif cmd == "mail-search":
            if not args:
                print("Usage: m365 mail-search <query>")
                sys.exit(1)
            msgs = client.search_mail(args[0])
            if not msgs:
                print("No matching messages found.")
            else:
                for m in msgs:
                    tag = "" if m['is_read'] else "[NEW] "
                    print(f"{tag}{m['date']} | ID: {m['id']} | {m['from_name']} <{m['from_address']}>: {m['subject']}")
                    if m.get('body'):
                        # Print first 200 chars of body for preview
                        preview = m['body'].replace('\n', ' ').replace('\r', '')[:200]
                        print(f"  Preview: {preview}")

        elif cmd == "mail-headers":
            if not args:
                print("Usage: m365 mail-headers <message_id>")
                sys.exit(1)
            import json
            headers = client.get_mail_headers(args[0])
            print(json.dumps(headers, indent=2, ensure_ascii=False))

        elif cmd == "mail-reply":
            if len(args) < 2:
                print("Usage: m365 mail-reply <message_id> <body>")
                sys.exit(1)
            print(client.reply_mail(args[0], args[1]))

        elif cmd == "mail-reply-all":
            if len(args) < 2:
                print("Usage: m365 mail-reply-all <message_id> <body>")
                sys.exit(1)
            print(client.reply_all_mail(args[0], args[1]))

        elif cmd == "mail-forward":
            if len(args) < 2:
                print("Usage: m365 mail-forward <message_id> <to> [body]")
                sys.exit(1)
            body = args[2] if len(args) > 2 else ""
            print(client.forward_mail(args[0], args[1], body))

        elif cmd == "mail-delete":
            if not args:
                print("Usage: m365 mail-delete <message_id>")
                sys.exit(1)
            print(client.delete_mail(args[0]))

        elif cmd == "mail-move":
            if len(args) < 2:
                print("Usage: m365 mail-move <message_id> <folder>")
                print("Well-known folders: inbox, sent, drafts, deleted, archive, junk")
                sys.exit(1)
            print(client.move_mail(args[0], args[1]))

        # ── Calendar ─────────────────────────────────────────────────────────

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
            import argparse
            parser = argparse.ArgumentParser(prog='m365 calendar-create')
            parser.add_argument('subject', help='Event title')
            parser.add_argument('start_iso', help='Start datetime (ISO 8601)')
            parser.add_argument('end_iso', help='End datetime (ISO 8601)')
            parser.add_argument('body', nargs='?', default='', help='Event description')
            parser.add_argument('--attendees', default='',
                                help='Required attendees, comma-separated emails')
            parser.add_argument('--optional-attendees', default='',
                                help='Optional attendees, comma-separated emails')
            parser.add_argument('--location', default='', help='Location or address')
            parser.add_argument('--private', action='store_true',
                                help='Mark event as private')
            parser.add_argument('--reminder', type=int, metavar='MINUTES',
                                help='Reminder N minutes before event')
            parser.add_argument('--attach', metavar='FILE', default=None,
                                help='Local file to attach to event')
            pargs = parser.parse_args(args)
            req_att = [e.strip() for e in pargs.attendees.split(',') if e.strip()]
            opt_att = [e.strip() for e in pargs.optional_attendees.split(',') if e.strip()]
            print(client.create_calendar_event(
                pargs.subject,
                pargs.start_iso,
                pargs.end_iso,
                body=pargs.body,
                location=pargs.location,
                required_attendees=req_att or None,
                optional_attendees=opt_att or None,
                is_private=pargs.private,
                reminder_minutes=pargs.reminder,
                attachment=pargs.attach,
            ))

        # ── Contacts ─────────────────────────────────────────────────────────

        elif cmd == "contacts-list":
            limit = int(args[0]) if args else 100
            contacts = client.list_contacts(limit=limit)
            if not contacts:
                print("No contacts found.")
            else:
                header = f"{'First':<15} {'Last':<15} {'Work Email':<30} {'Home Email':<25} {'Work Phone':<18} {'Home Phone':<15} Mobile"
                print(header)
                print("-" * len(header))
                for c in contacts:
                    print(
                        f"{c['first_name']:<15} {c['last_name']:<15} "
                        f"{c['work_email']:<30} {c['home_email']:<25} "
                        f"{c['work_phone']:<18} {c['home_phone']:<15} {c['mobile_phone']}"
                    )

        elif cmd == "contacts-get":
            if not args:
                print("Usage: m365 contacts-get <name>")
                sys.exit(1)
            import json
            contact = client.get_contact(args[0])
            if not contact:
                print(f"No contact found matching '{args[0]}'.")
            else:
                print(json.dumps(contact, indent=2, ensure_ascii=False))

        elif cmd == "contacts-create":
            import argparse
            parser = argparse.ArgumentParser(prog='m365 contacts-create')
            parser.add_argument('given_name', help='First name')
            parser.add_argument('surname', help='Last name')
            parser.add_argument('email', nargs='?', default=None,
                                help='Work email address (positional, optional)')
            parser.add_argument('phone', nargs='?', default=None,
                                help='Work phone (positional, optional)')
            parser.add_argument('--home-email', default=None, metavar='EMAIL')
            parser.add_argument('--home-phone', default=None, metavar='PHONE')
            parser.add_argument('--mobile-phone', default=None, metavar='PHONE')
            parser.add_argument('--work-street', default=None)
            parser.add_argument('--work-city', default=None)
            parser.add_argument('--work-state', default=None)
            parser.add_argument('--work-zip', default=None)
            parser.add_argument('--work-country', default=None)
            parser.add_argument('--home-street', default=None)
            parser.add_argument('--home-city', default=None)
            parser.add_argument('--home-state', default=None)
            parser.add_argument('--home-zip', default=None)
            parser.add_argument('--home-country', default=None)
            parser.add_argument('--birthday', default=None,
                                metavar='YYYY-MM-DDT00:00:00Z',
                                help='Birthday in ISO 8601 format')
            parser.add_argument('--anniversary', default=None,
                                metavar='YYYY-MM-DDT00:00:00Z')
            parser.add_argument('--website', default=None, help='Personal website URL')
            parser.add_argument('--work-website', default=None, help='Work website URL')
            parser.add_argument('--spouse', default=None, help="Spouse's name")
            parser.add_argument('--notes', default=None,
                                help='Free-text notes (use for hobbies, zodiac, children etc.)')
            pargs = parser.parse_args(args)
            print(client.create_contact(
                pargs.given_name, pargs.surname,
                email=pargs.email, phone=pargs.phone,
                home_email=pargs.home_email,
                home_phone=pargs.home_phone,
                mobile_phone=pargs.mobile_phone,
                work_street=pargs.work_street, work_city=pargs.work_city,
                work_state=pargs.work_state, work_zip=pargs.work_zip,
                work_country=pargs.work_country,
                home_street=pargs.home_street, home_city=pargs.home_city,
                home_state=pargs.home_state, home_zip=pargs.home_zip,
                home_country=pargs.home_country,
                birthday=pargs.birthday, anniversary=pargs.anniversary,
                website=pargs.website, work_website=pargs.work_website,
                spouse=pargs.spouse, notes=pargs.notes,
            ))

        elif cmd == "contacts-photo-set":
            if len(args) < 2:
                print("Usage: m365 contacts-photo-set <contact_id> <photo_path>")
                sys.exit(1)
            print(client.set_contact_photo(args[0], args[1]))

        elif cmd == "contacts-photo-delete":
            if not args:
                print("Usage: m365 contacts-photo-delete <contact_id>")
                sys.exit(1)
            print(client.delete_contact_photo(args[0]))

        elif cmd == "contacts-photo-get":
            if len(args) < 2:
                print("Usage: m365 contacts-photo-get <contact_id> <save_path>")
                sys.exit(1)
            print(client.get_contact_photo(args[0], args[1]))

        # ── OneDrive ─────────────────────────────────────────────────────────

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

        # ── OneNote ──────────────────────────────────────────────────────────

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

        # ── Microsoft ToDo ────────────────────────────────────────────────────

        elif cmd == "todo-list-lists":
            lists = client.todo_list_task_lists()
            if not lists:
                print("No task lists found.")
            else:
                print(f"{'ID':<50} Name")
                print("-" * 70)
                for lst in lists:
                    shared = " [shared]" if lst.get('is_shared') else ""
                    print(f"{lst['id']:<50} {lst['name']}{shared}")

        elif cmd == "todo-create-list":
            if not args:
                print("Usage: m365 todo-create-list <name>")
                sys.exit(1)
            print(client.todo_create_task_list(args[0]))

        elif cmd == "todo-rename-list":
            if len(args) < 2:
                print("Usage: m365 todo-rename-list <list_id> <new_name>")
                sys.exit(1)
            print(client.todo_rename_task_list(args[0], args[1]))

        elif cmd == "todo-delete-list":
            if not args:
                print("Usage: m365 todo-delete-list <list_id>")
                sys.exit(1)
            print(client.todo_delete_task_list(args[0]))

        elif cmd == "todo-list-tasks":
            import argparse
            parser = argparse.ArgumentParser(prog='m365 todo-list-tasks')
            parser.add_argument('list_id', help='Task list ID')
            parser.add_argument('--due-after', default=None,
                                metavar='YYYY-MM-DD',
                                help='Only show tasks due on or after this date')
            pargs = parser.parse_args(args)
            tasks = client.todo_list_tasks(pargs.list_id, due_after=pargs.due_after)
            _print_tasks(tasks)

        elif cmd == "todo-all-tasks":
            import argparse
            parser = argparse.ArgumentParser(prog='m365 todo-all-tasks')
            parser.add_argument('--due-after', default=None, metavar='YYYY-MM-DD')
            pargs = parser.parse_args(args)
            tasks = client.todo_get_all_tasks(due_after=pargs.due_after)
            _print_tasks(tasks, show_list=True)

        elif cmd == "todo-create-task":
            import argparse
            parser = argparse.ArgumentParser(prog='m365 todo-create-task')
            parser.add_argument('list_id', help='Task list ID')
            parser.add_argument('title', help='Task title')
            parser.add_argument('--note', default=None, help='Task note / description')
            parser.add_argument('--due', default=None,
                                metavar='YYYY-MM-DDTHH:MM:SS',
                                help='Due date (ISO 8601)')
            parser.add_argument('--reminder', default=None,
                                metavar='YYYY-MM-DDTHH:MM:SS',
                                help='Reminder datetime (ISO 8601)')
            pargs = parser.parse_args(args)
            print(client.todo_create_task(
                pargs.list_id, pargs.title,
                note=pargs.note,
                due_date=pargs.due,
                reminder_datetime=pargs.reminder,
            ))

        elif cmd == "todo-update-task":
            import argparse
            parser = argparse.ArgumentParser(prog='m365 todo-update-task')
            parser.add_argument('list_id', help='Task list ID')
            parser.add_argument('task_id', help='Task ID')
            parser.add_argument('--title', default=None)
            parser.add_argument('--note', default=None)
            parser.add_argument('--due', default=None, metavar='YYYY-MM-DDTHH:MM:SS')
            parser.add_argument('--reminder', default=None, metavar='YYYY-MM-DDTHH:MM:SS')
            pargs = parser.parse_args(args)
            print(client.todo_update_task(
                pargs.list_id, pargs.task_id,
                title=pargs.title,
                note=pargs.note,
                due_date=pargs.due,
                reminder_datetime=pargs.reminder,
            ))

        elif cmd == "todo-complete-task":
            if len(args) < 2:
                print("Usage: m365 todo-complete-task <list_id> <task_id>")
                sys.exit(1)
            print(client.todo_complete_task(args[0], args[1]))

        elif cmd == "todo-add-step":
            if len(args) < 3:
                print("Usage: m365 todo-add-step <list_id> <task_id> <step_title>")
                sys.exit(1)
            print(client.todo_add_step(args[0], args[1], args[2]))

        elif cmd == "todo-complete-step":
            if len(args) < 3:
                print("Usage: m365 todo-complete-step <list_id> <task_id> <step_id>")
                sys.exit(1)
            print(client.todo_complete_step(args[0], args[1], args[2]))

        elif cmd == "todo-move-task":
            if len(args) < 3:
                print("Usage: m365 todo-move-task <from_list_id> <task_id> <to_list_id>")
                sys.exit(1)
            print(client.todo_move_task(args[0], args[1], args[2]))

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


def _print_tasks(tasks, show_list=False):
    """Pretty-print a list of task dicts."""
    if not tasks:
        print("No tasks found.")
        return
    for t in tasks:
        done = "✓" if t.get('is_done') else "○"
        due = f" (due: {t['due'][:10]})" if t.get('due') else ""
        list_info = f" [{t.get('list_name', '')}]" if show_list else ""
        print(f"{done} {t['title']}{due}{list_info}")
        print(f"  ID: {t['id']}")
        if t.get('note'):
            preview = t['note'].replace('\n', ' ')[:100]
            print(f"  Note: {preview}")
        for step in t.get('steps', []):
            s_done = "✓" if step['is_done'] else "○"
            print(f"    {s_done} {step['title']}  (step ID: {step['id']})")


if __name__ == "__main__":
    main()
PYEOF

# 9. Create SKILL.md for OpenClaw agent auto-discovery
cat > "${INSTALL_DIR}/SKILL.md" << EOF
# SKILL: m365-graph  (v0.2.0)

OpenClaw skill for unattended Microsoft 365 access via Graph API.

## Executable
m365

## Description
Enables OpenClaw agents to send/read mail, manage calendar events, manage
contacts (incl. photos), access OneDrive files, create OneNote pages, edit
Excel/Word/PowerPoint documents, and fully manage Microsoft ToDo task lists
and tasks.  Uses client-credentials (daemon) flow – no browser required after
initial setup.

## Configuration
Credentials stored in: ${CONFIG_FILE}

Required Graph API permissions (Application):
  Mail.ReadWrite.All, Mail.Send, Calendars.ReadWrite.All,
  Files.ReadWrite.All, Contacts.ReadWrite, Tasks.ReadWrite.All

Required Graph API permissions (Delegated – OneNote only):
  Notes.ReadWrite.All, offline_access

## Commands

### Mail

| Command | Description | Example |
|---------|-------------|---------|
| \`mail-list [N]\` | List N recent inbox messages with IDs (default 20) | \`m365 mail-list 50\` |
| \`send-mail <to> <subject> <body> [opts]\` | Send an email. Opts: --cc, --bcc, --importance High\|Normal\|Low, --sensitivity Normal\|Personal\|Private\|Confidential, --attach /path, --delivery-receipt, --read-receipt | \`m365 send-mail alice@acme.com "Hello" "Hi there" --cc bob@acme.com --attach report.pdf\` |
| \`mail-search <query>\` | Search inbox by subject keyword; returns IDs, sender, preview | \`m365 mail-search "Invoice"\` |
| \`mail-headers <message_id>\` | Show full headers/metadata of a message (JSON) | \`m365 mail-headers AAMk…\` |
| \`mail-reply <message_id> <body>\` | Reply to the sender only | \`m365 mail-reply AAMk… "Thanks!"\` |
| \`mail-reply-all <message_id> <body>\` | Reply to all recipients | \`m365 mail-reply-all AAMk… "Noted by all."\` |
| \`mail-forward <message_id> <to> [body]\` | Forward message to a new recipient | \`m365 mail-forward AAMk… ceo@acme.com "FYI"\` |
| \`mail-delete <message_id>\` | Permanently delete a message | \`m365 mail-delete AAMk…\` |
| \`mail-move <message_id> <folder>\` | Move to folder (inbox, sent, drafts, deleted, archive, junk or folder ID) | \`m365 mail-move AAMk… archive\` |

### Calendar

| Command | Description | Example |
|---------|-------------|---------|
| \`calendar-list [days]\` | List upcoming events (default 7 days) | \`m365 calendar-list 14\` |
| \`calendar-create <subject> <start> <end> [body] [opts]\` | Create event. Opts: --attendees email,…, --optional-attendees email,…, --location "Place", --private, --reminder <minutes>, --attach /path | \`m365 calendar-create "Sync" 2026-04-15T10:00:00 2026-04-15T11:00:00 "Agenda" --attendees bob@acme.com --location "Room 3" --reminder 15\` |

### Contacts

| Command | Description | Example |
|---------|-------------|---------|
| \`contacts-list [N]\` | List N contacts as a table: First, Last, Work Mail, Home Mail, Work Phone, Home Phone, Mobile | \`m365 contacts-list\` |
| \`contacts-get <name>\` | Show all fields for the first contact whose name matches (JSON) | \`m365 contacts-get "Jane Doe"\` |
| \`contacts-create <first> <last> [email] [phone] [opts]\` | Create contact. Opts: --home-email, --home-phone, --mobile-phone, --work-street/city/zip/state/country, --home-street/city/zip/state/country, --birthday, --anniversary, --website, --work-website, --spouse, --notes | \`m365 contacts-create Jane Doe jane@acme.com +1-555-0100 --home-email jane@home.com --birthday 1985-06-15T00:00:00Z\` |
| \`contacts-photo-set <contact_id> <photo>\` | Upload a JPEG profile photo for a contact | \`m365 contacts-photo-set AAMk… /tmp/jane.jpg\` |
| \`contacts-photo-delete <contact_id>\` | Remove the profile photo for a contact | \`m365 contacts-photo-delete AAMk…\` |
| \`contacts-photo-get <contact_id> <save_path>\` | Download a contact's profile photo | \`m365 contacts-photo-get AAMk… /tmp/jane.jpg\` |

### OneDrive / Files

| Command | Description | Example |
|---------|-------------|---------|
| \`onedrive-list [folder]\` | List OneDrive folder (default /) | \`m365 onedrive-list /Documents\` |
| \`upload <local> <remote>\` | Upload a local file to OneDrive | \`m365 upload ./report.pdf /Reports/Q1.pdf\` |
| \`download <remote> <local>\` | Download a file from OneDrive | \`m365 download /Reports/Q1.pdf ./Q1.pdf\` |

### OneNote

| Command | Description | Example |
|---------|-------------|---------|
| \`onenote-create <nb> <section> <title> <html>\` | Create a OneNote page (delegated permission required) | \`m365 onenote-create "WorkNotes" "April" "Meeting" "<h1>Notes</h1>"\` |

### Excel / Word / PowerPoint

| Command | Description | Example |
|---------|-------------|---------|
| \`excel-update <path> <sheet> <range> [val …]\` | Update cells in an OneDrive Excel file | \`m365 excel-update report.xlsx Sheet1 A1:B2 Q1 Q2 100 200\` |
| \`word-update <path> key=val …\` | Replace placeholders in a Word document | \`m365 word-update offer.docx "{{Name}}=Alice" "{{Date}}=April 2026"\` |
| \`ppt-update <path> <slide> key=val …\` | Replace text in a PowerPoint slide (0-indexed) | \`m365 ppt-update deck.pptx 0 "{{Title}}=2026 Strategy"\` |

### Microsoft ToDo

| Command | Description | Example |
|---------|-------------|---------|
| \`todo-list-lists\` | List all task lists with IDs | \`m365 todo-list-lists\` |
| \`todo-create-list <name>\` | Create a new task list | \`m365 todo-create-list "Project Alpha"\` |
| \`todo-rename-list <list_id> <new_name>\` | Rename a task list | \`m365 todo-rename-list AAMk… "Project Beta"\` |
| \`todo-delete-list <list_id>\` | Delete a task list and all tasks | \`m365 todo-delete-list AAMk…\` |
| \`todo-list-tasks <list_id> [--due-after YYYY-MM-DD]\` | List tasks in a list (optional date filter) | \`m365 todo-list-tasks AAMk… --due-after 2026-04-01\` |
| \`todo-all-tasks [--due-after YYYY-MM-DD]\` | List all tasks across all lists | \`m365 todo-all-tasks --due-after 2026-04-01\` |
| \`todo-create-task <list_id> <title> [opts]\` | Create a task. Opts: --note, --due YYYY-MM-DDTHH:MM:SS, --reminder YYYY-MM-DDTHH:MM:SS | \`m365 todo-create-task AAMk… "Review budget" --due 2026-04-20T09:00:00 --note "Check Q2 numbers"\` |
| \`todo-update-task <list_id> <task_id> [opts]\` | Update task fields (--title, --note, --due, --reminder) | \`m365 todo-update-task AAMk… BBMk… --due 2026-04-25T09:00:00\` |
| \`todo-complete-task <list_id> <task_id>\` | Mark a task as completed | \`m365 todo-complete-task AAMk… BBMk…\` |
| \`todo-add-step <list_id> <task_id> <step>\` | Add a checklist step (subtask) | \`m365 todo-add-step AAMk… BBMk… "Draft email"\` |
| \`todo-complete-step <list_id> <task_id> <step_id>\` | Mark a step as completed | \`m365 todo-complete-step AAMk… BBMk… CCMk…\` |
| \`todo-move-task <from_list_id> <task_id> <to_list_id>\` | Move a task to a different list | \`m365 todo-move-task AAMk… BBMk… DDMk…\` |

## Notes
- **Mail**: Exchange Online Application Access Policy required in addition to
  Mail.ReadWrite.All. Ask your Exchange admin to run New-ApplicationAccessPolicy.
  See https://aka.ms/graph-app-access-policy
- **ToDo**: Requires the \`Tasks.ReadWrite.All\` application permission (grant
  admin consent in Entra ID → App registrations → API permissions).
- **OneNote**: Requires delegated permission Notes.ReadWrite.All; application
  permissions do not support OneNote as of early 2025.
- **Excel**: The workbook must not be open in a desktop Office application.
- **Word / PowerPoint**: Uses download → local edit → re-upload workflow.
- **Contact fields**: Hobbies, zodiac sign, and children count are not Graph
  API contact properties. Store these in the --notes field as free text.
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
echo "Available commands (v0.2.0):"
echo ""
echo "── MAIL ─────────────────────────────────────────────────────────────────"
echo "  m365 mail-list [N]                   – list recent inbox messages (with IDs)"
echo "  m365 send-mail <to> <sub> <body>      – send email"
echo "       [--cc addr] [--bcc addr]         – CC / BCC recipients"
echo "       [--importance High|Normal|Low]   – set priority"
echo "       [--sensitivity Normal|Personal|Private|Confidential]"
echo "       [--attach /path] [--delivery-receipt] [--read-receipt]"
echo "  m365 mail-search <query>              – search inbox by subject"
echo "  m365 mail-headers <message_id>        – show message headers"
echo "  m365 mail-reply <id> <body>           – reply to sender"
echo "  m365 mail-reply-all <id> <body>       – reply to all"
echo "  m365 mail-forward <id> <to> [body]    – forward message"
echo "  m365 mail-delete <id>                 – delete message"
echo "  m365 mail-move <id> <folder>          – move to folder (inbox/sent/archive/…)"
echo ""
echo "── CALENDAR ─────────────────────────────────────────────────────────────"
echo "  m365 calendar-list [days]             – list upcoming events"
echo "  m365 calendar-create <subj> <start> <end> [body]"
echo "       [--attendees email,…] [--optional-attendees email,…]"
echo "       [--location Place] [--private] [--reminder minutes] [--attach file]"
echo ""
echo "── CONTACTS ─────────────────────────────────────────────────────────────"
echo "  m365 contacts-list [N]                – list contacts (table format)"
echo "  m365 contacts-get <name>              – show all fields for one contact"
echo "  m365 contacts-create <first> <last> [email] [phone]"
echo "       [--home-email e] [--home-phone p] [--mobile-phone p]"
echo "       [--work-street/city/zip/country] [--home-street/city/zip/country]"
echo "       [--birthday ISO] [--anniversary ISO] [--website URL]"
echo "       [--work-website URL] [--spouse name] [--notes text]"
echo "  m365 contacts-photo-set <id> <file>   – upload contact photo (JPEG)"
echo "  m365 contacts-photo-delete <id>       – delete contact photo"
echo "  m365 contacts-photo-get <id> <file>   – download contact photo"
echo ""
echo "── ONEDRIVE ─────────────────────────────────────────────────────────────"
echo "  m365 onedrive-list [folder]           – list OneDrive folder"
echo "  m365 upload <local> <remote>          – upload file"
echo "  m365 download <remote> <local>        – download file"
echo ""
echo "── ONENOTE ──────────────────────────────────────────────────────────────"
echo "  m365 onenote-create <nb> <sec> <title> <html>"
echo ""
echo "── EXCEL / WORD / POWERPOINT ────────────────────────────────────────────"
echo "  m365 excel-update <path> <sheet> <range> [val …]"
echo "  m365 word-update <path> key=val …"
echo "  m365 ppt-update <path> <slide> key=val …"
echo ""
echo "── MICROSOFT TODO ───────────────────────────────────────────────────────"
echo "  m365 todo-list-lists                  – list all task lists"
echo "  m365 todo-create-list <name>          – create task list"
echo "  m365 todo-rename-list <id> <name>     – rename task list"
echo "  m365 todo-delete-list <id>            – delete task list"
echo "  m365 todo-list-tasks <list_id> [--due-after YYYY-MM-DD]"
echo "  m365 todo-all-tasks [--due-after YYYY-MM-DD]"
echo "  m365 todo-create-task <list_id> <title> [--note t] [--due ISO] [--reminder ISO]"
echo "  m365 todo-update-task <list_id> <task_id> [--title t] [--note t] [--due ISO]"
echo "  m365 todo-complete-task <list_id> <task_id>"
echo "  m365 todo-add-step <list_id> <task_id> <step_title>"
echo "  m365 todo-complete-step <list_id> <task_id> <step_id>"
echo "  m365 todo-move-task <from_id> <task_id> <to_id>"
echo ""
echo "If authentication fails or no calendar appears:"
echo "  rm -rf ${OPENCLAW_DIR}/credentials/m365_token_cache.bin"
echo "  cd ${INSTALL_DIR}"
echo "  ${VENV_DIR}/bin/python -m m365_openclaw calendar-list"
echo ""
echo "Done!"
