"""
core.py – Microsoft 365 client for OpenClaw agents (v0.3.1).

Supports: Mail, Calendar, Contacts, OneDrive, OneNote, Excel, Word, PowerPoint,
          Microsoft ToDo tasks.
Uses client-credentials (daemon/application) flow via the O365 library + MSAL.

Important: When using application permissions (client_credentials flow), all
Graph API calls must target a specific user. Set M365_USER_EMAIL in .env so
that main_resource resolves /me/ to /users/<email>/ automatically.
"""

import base64
import os
import re
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

import msal
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

        # MSAL app used for direct Graph API calls (todo-*, etc.)
        # Using MSAL directly avoids reading O365 internal connection state,
        # which is unreliable and can corrupt the O365 session when retried.
        self._msal_app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret,
        )

        if not self.account.is_authenticated:
            print("Authenticating with Microsoft 365...", file=sys.stderr)
            self.account.authenticate(
                scopes=['https://graph.microsoft.com/.default']
            )

    # ── internal helpers ──────────────────────────────────────────────────────

    def _get_drive(self):
        storage = self.account.storage()
        return storage.get_default_drive(request_drive=True)

    def _access_token(self):
        """Return a valid access token using MSAL client credentials flow.

        Uses self._msal_app (ConfidentialClientApplication) directly instead
        of reading from the O365 connection's internal state.  The O365
        library's own session (account.schedule(), etc.) is kept intact
        because we never call account.authenticate() here again.
        """
        scope = ["https://graph.microsoft.com/.default"]

        # Try a silently cached token first (avoids a round-trip to MSFT)
        result = self._msal_app.acquire_token_silent(scope, account=None)
        if result and "access_token" in result:
            return result["access_token"]

        # Fetch a fresh token from Azure AD
        print("Refreshing Microsoft 365 token...", file=sys.stderr)
        result = self._msal_app.acquire_token_for_client(scopes=scope)
        if result and "access_token" in result:
            return result["access_token"]

        error = result.get("error", "unknown") if result else "no response"
        desc = result.get("error_description", "") if result else ""
        raise RuntimeError(
            f"Could not retrieve a valid access token.\n"
            f"Azure AD error: {error} – {desc}\n"
            "Please verify TENANT_ID, CLIENT_ID, and CLIENT_SECRET in .env"
        )

    def _graph_headers(self, content_type='application/json'):
        headers = {'Authorization': f'Bearer {self._access_token()}'}
        if content_type:
            headers['Content-Type'] = content_type
        return headers

    def _graph_get(self, url, params=None):
        resp = requests.get(url, headers=self._graph_headers(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_post(self, url, data=None, content_type='application/json', raw_data=None):
        headers = self._graph_headers(content_type=content_type)
        if raw_data is not None:
            resp = requests.post(url, headers=headers, data=raw_data, timeout=30)
        else:
            resp = requests.post(url, headers=headers, json=data, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_patch(self, url, data):
        resp = requests.patch(url, headers=self._graph_headers(), json=data, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _graph_delete(self, url):
        resp = requests.delete(url, headers=self._graph_headers(content_type=None), timeout=30)
        resp.raise_for_status()
        return {}

    def _graph_put(self, url, data, content_type='application/octet-stream'):
        headers = self._graph_headers(content_type=content_type)
        resp = requests.put(url, headers=headers, data=data, timeout=60)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def _base_url(self):
        return f"https://graph.microsoft.com/v1.0/users/{self.user}"

    # ── Mail – error helper ───────────────────────────────────────────────────

    @staticmethod
    def _raise_if_mail_403(exc):
        try:
            import requests
            if isinstance(exc, requests.exceptions.HTTPError) and exc.response.status_code == 403:
                raise PermissionError(
                    "Mail access denied (HTTP 403).\n"
                    "Make sure RBAC for Applications is configured via setup-exchange-policy.ps1\n"
                    "and admin consent is granted in Entra ID."
                ) from exc
        except ImportError:
            pass

    # ── Todo – error helper ───────────────────────────────────────────────────

    @staticmethod
    def _raise_if_todo_401(exc):
        try:
            import requests
            if isinstance(exc, requests.exceptions.HTTPError) and exc.response.status_code == 401:
                raise PermissionError(
                    "Microsoft To Do access denied (HTTP 401 Unauthorized).\n"
                    "Make sure the application permission 'Tasks.ReadWrite.All' is added\n"
                    "in Entra ID → App registrations → API permissions, and that\n"
                    "admin consent has been granted for your organization."
                ) from exc
        except ImportError:
            pass

    # ── Mail ──────────────────────────────────────────────────────────────────

    def send_mail(self, to_address, subject, body, cc=None, bcc=None,
                  sensitivity='Normal', importance='Normal', attachments=None,
                  request_delivery_receipt=False, request_read_receipt=False):
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

    def list_mail(self, limit=20, folder='inbox', unread_only=False, sort='new-old',
                  search_by_email=None, search_by_subject=None, search=None,
                  count_only=False):
        well_known = {
            'inbox': 'inbox',
            'sent': 'sentitems',
            'drafts': 'drafts',
            'deleted': 'deleteditems',
            'archive': 'archive',
            'junk': 'junkemail',
        }
        folder_key = well_known.get((folder or 'inbox').lower(), None)
        if folder_key:
            msg_url = f"{self._base_url()}/mailFolders/{folder_key}/messages"
        else:
            folders_data = self._graph_get(
                f"{self._base_url()}/mailFolders",
                params={'$filter': f"displayName eq '{folder}'"},
            )
            if folders_data.get('value'):
                fid = folders_data['value'][0]['id']
                msg_url = f"{self._base_url()}/mailFolders/{fid}/messages"
            else:
                msg_url = f"{self._base_url()}/mailFolders/inbox/messages"

        if count_only:
            try:
                count_headers = self._graph_headers(content_type=None)
                count_headers['ConsistencyLevel'] = 'eventual'
                resp = requests.get(
                    msg_url + '/$count', headers=count_headers, timeout=30
                )
                resp.raise_for_status()
                return {'count': int(resp.text)}
            except Exception as exc:
                self._raise_if_mail_403(exc)
                raise

        params = {
            '$top': min(int(limit), 500),
            '$select': 'id,subject,from,receivedDateTime,isRead',
        }
        if search:
            params['$search'] = f'"{search}"'
        else:
            params['$orderby'] = (
                'receivedDateTime desc' if sort != 'old-new' else 'receivedDateTime asc'
            )
            filters = []
            if unread_only:
                filters.append('isRead eq false')
            if search_by_email:
                filters.append(f"from/emailAddress/address eq '{search_by_email}'")
            if search_by_subject:
                filters.append(f"contains(subject,'{search_by_subject}')")
            if filters:
                params['$filter'] = ' and '.join(filters)

        try:
            data = self._graph_get(msg_url, params=params)
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        result = []
        for msg in data.get('value', []):
            from_addr = msg.get('from', {}).get('emailAddress', {})
            result.append({
                'id': msg.get('id', ''),
                'subject': msg.get('subject', '(no subject)'),
                'from': f"{from_addr.get('name', '')} <{from_addr.get('address', '')}>",
                'from_address': from_addr.get('address', ''),
                'date': msg.get('receivedDateTime', ''),
                'is_read': msg.get('isRead', False),
            })
        return result

    def read_mail(self, message_id, mark_as_read=False, mark_as_unread=False,
                  full_body=False, headers_only=False):
        select_fields = (
            'id,subject,from,toRecipients,ccRecipients,'
            'receivedDateTime,sentDateTime,importance,sensitivity,'
            'isRead,hasAttachments,body'
        )
        url = f"{self._base_url()}/messages/{message_id}?$select={select_fields}"
        try:
            msg = self._graph_get(url)
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise

        body_content = msg.get('body', {}).get('content', '')
        if headers_only:
            body_content = None
        elif not full_body:
            body_text = re.sub(r'<[^>]+>', ' ', body_content)
            body_text = ' '.join(body_text.split())
            body_content = body_text[:200] if len(body_text) > 200 else body_text

        result = {
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
        if body_content is not None:
            result['body'] = body_content

        if mark_as_read or mark_as_unread:
            patch_url = f"{self._base_url()}/messages/{message_id}"
            is_read = True if mark_as_read else False
            try:
                self._graph_patch(patch_url, {'isRead': is_read})
            except Exception as exc:
                self._raise_if_mail_403(exc)
                raise
            result['is_read'] = is_read

        return result

    def create_mail_folder(self, folder_name):
        url = f"{self._base_url()}/mailFolders"
        try:
            result = self._graph_post(url, {'displayName': folder_name})
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Mail folder '{folder_name}' created (ID: {result.get('id', '')})"

    def delete_mail_folder(self, folder_id):
        well_known = {
            'inbox': 'inbox', 'sent': 'sentitems', 'drafts': 'drafts',
            'deleted': 'deleteditems', 'archive': 'archive', 'junk': 'junkemail',
        }
        fid = well_known.get(folder_id.lower(), folder_id)
        url = f"{self._base_url()}/mailFolders/{fid}"
        try:
            self._graph_delete(url)
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Mail folder '{folder_id}' deleted"

    def rename_mail_folder(self, folder_id, new_name):
        well_known = {
            'inbox': 'inbox', 'sent': 'sentitems', 'drafts': 'drafts',
            'deleted': 'deleteditems', 'archive': 'archive', 'junk': 'junkemail',
        }
        fid = well_known.get(folder_id.lower(), folder_id)
        url = f"{self._base_url()}/mailFolders/{fid}"
        try:
            self._graph_patch(url, {'displayName': new_name})
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Mail folder renamed to '{new_name}'"

    def search_mail(self, subject_query, limit=10):
        url = (
            f"{self._base_url()}/mailFolders/inbox/messages"
            f"?$top={limit}"
            f"&$select=id,subject,from,receivedDateTime,isRead,body"
        )
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
        url = f"{self._base_url()}/messages/{message_id}/reply"
        try:
            self._graph_post(url, {'message': {}, 'comment': body})
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Reply sent for message {message_id}"

    def reply_all_mail(self, message_id, body):
        url = f"{self._base_url()}/messages/{message_id}/replyAll"
        try:
            self._graph_post(url, {'message': {}, 'comment': body})
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Reply-All sent for message {message_id}"

    def forward_mail(self, message_id, to_address, body=""):
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
        url = f"{self._base_url()}/messages/{message_id}"
        try:
            self._graph_delete(url)
        except Exception as exc:
            self._raise_if_mail_403(exc)
            raise
        return f"Message {message_id} deleted"

    def move_mail(self, message_id, destination_folder):
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
        event_data = {
            'subject': subject,
            'body': {'contentType': 'HTML', 'content': body},
            'start': {'dateTime': start_iso, 'timeZone': 'UTC'},
            'end': {'dateTime': end_iso, 'timeZone': 'UTC'},
            'sensitivity': 'private' if is_private else 'normal',
        }
        if location:
            event_data['location'] = {'displayName': location}

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

        if reminder_minutes is not None:
            event_data['isReminderOn'] = True
            event_data['reminderMinutesBeforeStart'] = int(reminder_minutes)

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

    def get_calendar_event(self, event_id):
        url = f"{self._base_url()}/calendar/events/{event_id}"
        result = self._graph_get(url)
        attendees = []
        for a in result.get('attendees', []):
            attendees.append({
                'name': a.get('emailAddress', {}).get('name', ''),
                'email': a.get('emailAddress', {}).get('address', ''),
                'type': a.get('type', ''),
                'status': a.get('status', {}).get('response', ''),
            })
        return {
            'id': result.get('id', ''),
            'subject': result.get('subject', '(no subject)'),
            'start': result.get('start', {}).get('dateTime', ''),
            'end': result.get('end', {}).get('dateTime', ''),
            'location': result.get('location', {}).get('displayName', ''),
            'body': result.get('body', {}).get('content', ''),
            'sensitivity': result.get('sensitivity', ''),
            'is_organizer': result.get('isOrganizer', False),
            'is_cancelled': result.get('isCancelled', False),
            'attendees': attendees,
        }

    def respond_calendar_event(self, event_id, response):
        """Response must be one of: accept, tentativelyAccept, decline."""
        url = f"{self._base_url()}/calendar/events/{event_id}/{response}"
        self._graph_post(url, {'sendResponse': True})
        return f"Response '{response}' sent for event {event_id}"

    def cancel_calendar_event(self, event_id, comment=''):
        url = f"{self._base_url()}/calendar/events/{event_id}/cancel"
        self._graph_post(url, {'comment': comment})
        return f"Event {event_id} cancelled"

    def delete_calendar_event(self, event_id):
        url = f"{self._base_url()}/calendar/events/{event_id}"
        self._graph_delete(url)
        return f"Event {event_id} deleted"

    # ── Contacts ──────────────────────────────────────────────────────────────

    def list_contacts(self, limit=100):
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

        resolved_work_phone = work_phone or phone
        if resolved_work_phone:
            contact_data['businessPhones'] = [resolved_work_phone.strip()]
        if home_phone:
            contact_data['homePhones'] = [home_phone.strip()]
        if mobile_phone:
            contact_data['mobilePhone'] = mobile_phone.strip()

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

        if birthday:
            contact_data['birthday'] = birthday
        if anniversary:
            contact_data['anniversary'] = anniversary
        if spouse:
            contact_data['spouseName'] = spouse
        if notes:
            contact_data['personalNotes'] = notes

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
        with open(photo_path, 'rb') as fh:
            photo_bytes = fh.read()
        url = f"{self._base_url()}/contacts/{contact_id}/photo/$value"
        ext = Path(photo_path).suffix.lower()
        content_type = 'image/jpeg' if ext in ('.jpg', '.jpeg') else 'image/png'
        self._graph_put(url, photo_bytes, content_type=content_type)
        return f"Photo set for contact {contact_id}"

    def delete_contact_photo(self, contact_id):
        url = f"{self._base_url()}/contacts/{contact_id}/photo/$value"
        self._graph_delete(url)
        return f"Photo deleted for contact {contact_id}"

    def get_contact_photo(self, contact_id, save_path):
        url = f"{self._base_url()}/contacts/{contact_id}/photo/$value"
        resp = requests.get(url, headers=self._graph_headers(content_type=None), timeout=30)
        resp.raise_for_status()
        with open(save_path, 'wb') as fh:
            fh.write(resp.content)
        return f"Photo saved to {save_path}"

    def get_contact_by_id(self, contact_id):
        url = f"{self._base_url()}/contacts/{contact_id}"
        try:
            c = self._graph_get(url)
        except Exception:
            return None
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

    def update_contact(self, contact_id, given_name=None, surname=None,
                       email_business=None, email_personal=None,
                       phone_mobile=None, phone_business=None, phone_home=None,
                       birthday=None, notes=None, company=None, job_title=None):
        data = {}
        if given_name is not None:
            data['givenName'] = given_name
        if surname is not None:
            data['surname'] = surname
        if company is not None:
            data['companyName'] = company
        if job_title is not None:
            data['jobTitle'] = job_title
        if phone_mobile is not None:
            data['mobilePhone'] = phone_mobile
        if phone_business is not None:
            data['businessPhones'] = [phone_business]
        if phone_home is not None:
            data['homePhones'] = [phone_home]
        if birthday is not None:
            data['birthday'] = birthday
        if notes is not None:
            data['personalNotes'] = notes

        if email_business is not None or email_personal is not None:
            current = self._graph_get(
                f"{self._base_url()}/contacts/{contact_id}?$select=emailAddresses"
            )
            emails = list(current.get('emailAddresses', []))
            if email_business is not None:
                found = False
                for e in emails:
                    if e.get('name', '').lower() in ('work', 'geschäft', 'business'):
                        e['address'] = email_business
                        found = True
                        break
                if not found:
                    emails.append({'name': 'Work', 'address': email_business})
            if email_personal is not None:
                found = False
                for e in emails:
                    if e.get('name', '').lower() in ('home', 'personal', 'privat'):
                        e['address'] = email_personal
                        found = True
                        break
                if not found:
                    emails.append({'name': 'Home', 'address': email_personal})
            data['emailAddresses'] = emails

        if not data:
            return "No fields to update."
        url = f"{self._base_url()}/contacts/{contact_id}"
        self._graph_patch(url, data)
        return f"Contact {contact_id} updated"

    def list_contact_folders(self):
        url = f"{self._base_url()}/contactFolders"
        data = self._graph_get(url)
        return [
            {'id': f.get('id', ''), 'name': f.get('displayName', '')}
            for f in data.get('value', [])
        ]

    def create_contact_folder(self, name):
        url = f"{self._base_url()}/contactFolders"
        result = self._graph_post(url, {'displayName': name})
        return f"Contact folder '{name}' created (ID: {result.get('id', '')})"

    def delete_contact_folder(self, folder_id):
        url = f"{self._base_url()}/contactFolders/{folder_id}"
        self._graph_delete(url)
        return f"Contact folder {folder_id} deleted"

    # ── OneDrive ─────────────────────────────────────────────────────────────

    def onedrive_list(self, folder_path="/"):
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
        drive = self._get_drive()
        item = drive.get_item_by_path(remote_path)
        local = Path(local_path)
        item.download(to_path=str(local.parent), name=local.name)
        return f"Downloaded to {local_path}"

    def onedrive_delete(self, remote_path):
        drive = self._get_drive()
        item = drive.get_item_by_path(remote_path)
        item.delete()
        return f"Deleted {remote_path}"

    def onedrive_move(self, old_path, new_path):
        """Move or rename a file/folder on OneDrive."""
        drive = self._get_drive()
        item = drive.get_item_by_path(old_path)
        drive_id = drive.object_id
        item_id = item.object_id
        new = Path(new_path)
        new_parent = str(new.parent)
        if new_parent in ('/', '.', ''):
            parent_item = drive.get_root_folder()
        else:
            parent_item = drive.get_item_by_path(new_parent)
        parent_id = parent_item.object_id
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
        self._graph_patch(url, {
            'name': new.name,
            'parentReference': {'id': parent_id},
        })
        return f"Moved {old_path} to {new_path}"

    def onedrive_share(self, remote_path, anyone=False, edit=False):
        drive = self._get_drive()
        item = drive.get_item_by_path(remote_path)
        drive_id = drive.object_id
        item_id = item.object_id
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/createLink"
        data = {
            'type': 'edit' if edit else 'view',
            'scope': 'anonymous' if anyone else 'organization',
        }
        result = self._graph_post(url, data)
        link = result.get('link', {}).get('webUrl', '')
        return f"Share link: {link}"

    # ── User ──────────────────────────────────────────────────────────────────

    def get_user(self):
        url = f"{self._base_url()}"
        data = self._graph_get(url)
        return {
            'id': data.get('id', ''),
            'display_name': data.get('displayName', ''),
            'given_name': data.get('givenName', ''),
            'surname': data.get('surname', ''),
            'email': data.get('mail', '') or data.get('userPrincipalName', ''),
            'job_title': data.get('jobTitle', ''),
            'department': data.get('department', ''),
            'mobile_phone': data.get('mobilePhone', ''),
            'office_location': data.get('officeLocation', ''),
        }

    def update_user(self, display_name=None, given_name=None, surname=None,
                    mobile_phone=None, job_title=None, department=None,
                    office_location=None):
        data = {}
        if display_name is not None:
            data['displayName'] = display_name
        if given_name is not None:
            data['givenName'] = given_name
        if surname is not None:
            data['surname'] = surname
        if mobile_phone is not None:
            data['mobilePhone'] = mobile_phone
        if job_title is not None:
            data['jobTitle'] = job_title
        if department is not None:
            data['department'] = department
        if office_location is not None:
            data['officeLocation'] = office_location
        if not data:
            return "No fields to update."
        url = f"{self._base_url()}"
        self._graph_patch(url, data)
        return "User profile updated"

    # ── OneNote ───────────────────────────────────────────────────────────────

    def onenote_create_page(self, notebook_name, section_name, title, html_content):
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
        url = f"{self._base_url()}/todo/lists"
        try:
            data = self._graph_get(url)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
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
        url = f"{self._base_url()}/todo/lists"
        try:
            result = self._graph_post(url, {'displayName': name})
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Task list '{name}' created (ID: {result.get('id', '')})"

    def todo_rename_task_list(self, list_id, new_name):
        url = f"{self._base_url()}/todo/lists/{list_id}"
        try:
            self._graph_patch(url, {'displayName': new_name})
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Task list {list_id} renamed to '{new_name}'"

    def todo_delete_task_list(self, list_id):
        url = f"{self._base_url()}/todo/lists/{list_id}"
        try:
            self._graph_delete(url)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Task list {list_id} deleted"

    def todo_list_tasks(self, list_id, due_after=None, due_before=None):
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks?$top=100&$expand=checklistItems"
        try:
            data = self._graph_get(url)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        tasks = self._format_tasks(data.get('value', []))
        if due_after:
            tasks = [t for t in tasks if t.get('due') and t['due'][:10] >= due_after]
        if due_before:
            tasks = [t for t in tasks if t.get('due') and t['due'][:10] <= due_before]
        return tasks

    def todo_get_all_tasks(self, due_after=None, due_before=None):
        try:
            lists_data = self._graph_get(f"{self._base_url()}/todo/lists")
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        all_tasks = []
        for lst in lists_data.get('value', []):
            list_id = lst['id']
            list_name = lst.get('displayName', '')
            tasks = self.todo_list_tasks(list_id, due_after=due_after, due_before=due_before)
            for task in tasks:
                task['list_name'] = list_name
                task['list_id'] = list_id
            all_tasks.extend(tasks)
        return all_tasks

    def todo_get_tasks_today(self):
        from datetime import date
        today = date.today().isoformat()
        return self.todo_get_all_tasks(due_after=today, due_before=today)

    @staticmethod
    def _format_tasks(tasks):
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
        task_data = {'title': title}
        if note:
            task_data['body'] = {'content': note, 'contentType': 'text'}
        if due_date:
            task_data['dueDateTime'] = {'dateTime': due_date, 'timeZone': 'UTC'}
        if reminder_datetime:
            task_data['reminderDateTime'] = {'dateTime': reminder_datetime, 'timeZone': 'UTC'}
            task_data['isReminderOn'] = True
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks"
        try:
            result = self._graph_post(url, task_data)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Task '{title}' created (ID: {result.get('id', '')})"

    def todo_update_task(self, list_id, task_id, title=None, note=None,
                          due_date=None, reminder_datetime=None):
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
        try:
            self._graph_patch(url, task_data)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Task {task_id} updated"

    def todo_complete_task(self, list_id, task_id):
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks/{task_id}"
        try:
            self._graph_patch(url, {'status': 'completed'})
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Task {task_id} marked as completed"

    def todo_add_step(self, list_id, task_id, step_title):
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks/{task_id}/checklistItems"
        try:
            result = self._graph_post(url, {'displayName': step_title, 'isChecked': False})
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Step '{step_title}' added (ID: {result.get('id', '')})"

    def todo_complete_step(self, list_id, task_id, step_id):
        url = (
            f"{self._base_url()}/todo/lists/{list_id}"
            f"/tasks/{task_id}/checklistItems/{step_id}"
        )
        try:
            self._graph_patch(url, {'isChecked': True})
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Step {step_id} marked as completed"

    def todo_move_task(self, from_list_id, task_id, to_list_id):
        src_url = (
            f"{self._base_url()}/todo/lists/{from_list_id}"
            f"/tasks/{task_id}?$expand=checklistItems"
        )
        try:
            task = self._graph_get(src_url)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise

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
        try:
            new_task = self._graph_post(dst_url, task_data)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        new_task_id = new_task.get('id', '')

        for item in (task.get('checklistItems') or []):
            step_url = (
                f"{self._base_url()}/todo/lists/{to_list_id}"
                f"/tasks/{new_task_id}/checklistItems"
            )
            self._graph_post(step_url, {
                'displayName': item.get('displayName', ''),
                'isChecked': item.get('isChecked', False),
            })

        self._graph_delete(
            f"{self._base_url()}/todo/lists/{from_list_id}/tasks/{task_id}"
        )
        return (
            f"Task moved from list {from_list_id} to {to_list_id} "
            f"(new ID: {new_task_id})"
        )


    def todo_get_task(self, list_id, task_id):
        url = (
            f"{self._base_url()}/todo/lists/{list_id}"
            f"/tasks/{task_id}?$expand=checklistItems"
        )
        try:
            data = self._graph_get(url)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        tasks = self._format_tasks([data])
        return tasks[0] if tasks else None

    def todo_delete_task(self, list_id, task_id):
        url = f"{self._base_url()}/todo/lists/{list_id}/tasks/{task_id}"
        try:
            self._graph_delete(url)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        return f"Task {task_id} deleted"

    def todo_get_default_list_id(self):
        """Return the ID of the default task list using wellKnownListName, or the first list."""
        url = f"{self._base_url()}/todo/lists?$select=id,displayName,wellKnownListName"
        try:
            data = self._graph_get(url)
        except Exception as exc:
            self._raise_if_todo_401(exc)
            raise
        lists = data.get('value', [])
        if not lists:
            return None
        for lst in lists:
            if lst.get('wellKnownListName') == 'tasks':
                return lst.get('id', '')
        return lists[0].get('id', '')


if __name__ == "__main__":
    client = M365Client()
    print("M365Client initialized successfully")
