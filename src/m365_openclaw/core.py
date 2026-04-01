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
        # Escape single quotes in the subject for safe OData filter interpolation.
        safe_subject = subject.replace("'", "''")
        url = (
            f"{GRAPH_BASE}/users/{self.user}/messages"
            f"?$filter=contains(subject,'{safe_subject}')"
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
            cutoff = datetime.fromisoformat(due_before).astimezone(timezone.utc)
            tasks = [t for t in tasks if self._task_due(t) and self._task_due(t) <= cutoff]
        if due_after:
            floor = datetime.fromisoformat(due_after).astimezone(timezone.utc)
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
            due_date=payload.get('dueDateTime', {}).get('dateTime', '').split('T')[0] if payload.get('dueDateTime', {}).get('dateTime') else None,
        )
        self._graph_delete(src_url)
        return created
