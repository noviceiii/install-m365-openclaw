"""
__main__.py – CLI entry point for the OpenClaw M365 skill (v0.5.0).

Invoked as:  python -m m365_openclaw <command> [arguments...]
Or via the  m365  wrapper script placed in ~/.local/bin.
"""

import argparse
import json
import sys

USAGE = """\
OpenClaw M365 CLI v0.5.0 – Microsoft 365 for agents (delegated / device-code auth)

━━━ AUTH ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  auth-login
      Force a new device-code login (opens https://microsoft.com/devicelogin).
      Use this for initial setup or after token expiry.

━━━ MAIL ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  mail-list [N]
      List N recent messages (default: 20, max: 500).
      --sort new-old|old-new   Sort order (default: new-old)
      --unread                 Only unread messages
      --folder NAME            Folder name (default: inbox). Well-known:
                               inbox, sent, drafts, deleted, archive, junk
      --search-by-email ADDR   Filter by sender e-mail address
      --search-by-subject TEXT Filter by subject (contains)
      --search TEXT            Full-text search
      --count                  Print count only

  mail-send
      Send an e-mail.
      --to ADDR[,...]   Recipient(s), required
      --subject TEXT    Subject, required
      --body TEXT       Body text, required
      --html            Send body as HTML (contentType: HTML)
      --text            Send body as plain text (contentType: Text, default)
      --cc ADDR[,...]   Carbon copy
      --bcc ADDR[,...]  Blind carbon copy
      --attach FILE     Attach a local file (repeat for multiple)
      --priority High|Normal|Low   Priority (default: Normal)

  mail-read <message_id>
      Show details of a message (body truncated to 200 chars by default).
      --full-body      Show complete body
      --read-header    Show headers only (no body)
      --mark-read      Mark message as read
      --mark-unread    Mark message as unread

  mail-reply <message_id>
      Reply to a message.
      --body TEXT      Reply body, required
      --reply-all      Reply to all recipients

  mail-forward <message_id>
      Forward a message.
      --to ADDR        Recipient, required
      --body TEXT      Optional comment

  mail-handle <message_id>
      Delete, archive or move a message.
      --delete         Permanently delete
      --archive        Move to archive folder
      --move FOLDER    Move to named folder

  mailbox-handle
      Manage mail folders.
      --folder-list           List all mail folders (returns ID and Name)
      --folder-create NAME    Create a new folder
      --folder-delete NAME    Delete a folder (by name or ID)
      --folder-rename ID NAME Rename a folder

━━━ CALENDAR ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  calendar-list [N]
      List upcoming events (N = days into the future, default: 7).

  calendar-create
      Create a calendar event.
      --subject TEXT   Event title, required
      --start ISO      Start datetime (ISO 8601), required
      --end ISO        End datetime (ISO 8601), required
      --body TEXT      Description
      --location TEXT  Location
      --required EMAIL Required attendee(s), comma-separated
      --optional EMAIL Optional attendee(s), comma-separated
      --private        Mark as private
      --reminder N     Reminder N minutes before event
      --attach FILE    Attach a local file

  calendar-read <event_id>
      Show details of a calendar event.
      --participants   List participants and their status

  calendar-handle <event_id>
      Respond to or manage a calendar event.
      --confirm-accept    Accept the invitation
      --confirm-tentative Tentatively accept
      --confirm-deny      Decline the invitation
      --cancel            Cancel the event (organizer only)
      --cancel-no-info    Cancel without sending cancellation notice
      --delete            Permanently delete the event

━━━ CONTACTS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  contact-list [N]
      List N contacts (default: 100) as a table.
      --sort by-last|by-first   Sort order
      --count                   Print count only

  contact-read <contact_id>
      Show all fields of a contact by ID.

  contact-edit <contact_id>
      Update fields of an existing contact.
      --given-name TEXT        First name
      --surname TEXT           Last name
      --email-business EMAIL   Business e-mail
      --email-personal EMAIL   Personal e-mail
      --phone-mobile PHONE     Mobile phone
      --phone-business PHONE   Business phone
      --phone-home PHONE       Home phone
      --birthday YYYY-MM-DDT00:00:00Z
      --notes TEXT             Personal notes
      --company TEXT           Company name
      --job-title TEXT         Job title

  contact-create <first> <last>
      Create a new contact.
      --email-business EMAIL   Business e-mail
      --email-personal EMAIL   Personal e-mail
      --phone-mobile PHONE
      --phone-business PHONE
      --phone-home PHONE
      --birthday YYYY-MM-DDT00:00:00Z
      --notes TEXT

  contact-photo <contact_id>
      Manage a contact's profile photo.
      --upload FILE   Upload photo (JPEG/PNG)
      --delete        Delete photo
      --download PATH Save photo to local file

  contact-delete <contact_id>
      Delete a contact from the address book by ID.

  contactlist-list
      Show all contact folders / contact lists.

  contactlist-create <name>
      Create a new contact folder.

  contactlist-delete <folder_id>
      Delete a contact folder.

━━━ TASKS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  task-list [N]
      List tasks (default: 50).
      --list-id ID   Restrict to a specific task list
      --status STATUS   Filter: notStarted|inProgress|completed
      --count        Print count only

  task-create
      Create a new task.
      --title TEXT   Task title, required
      --list-id ID   Task list ID (defaults to first list)
      --due YYYY-MM-DDTHH:MM:SS   Due date
      --body TEXT    Notes / description
      --priority low|normal|high

  task-read <task_id>
      Show details of a task.
      --list-id ID   Task list ID (required if task ID is ambiguous)

  task-edit <task_id>
      Edit a task.
      --list-id ID               Task list ID
      --title TEXT
      --status notStarted|inProgress|completed
      --due YYYY-MM-DDTHH:MM:SS
      --complete                 Mark as completed
      --checklist-add TEXT       Add a checklist item
      --checklist-complete ID    Mark a checklist item as completed
      --checklist-delete ID      Delete a checklist item

  task-delete <task_id>
      Delete a task.
      --list-id ID   Task list ID

  tasklist-list
      List all task lists with their IDs.

  tasklist-create <name>
      Create a new task list.

━━━ USER ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  user-read
      Show the own user profile.

  user-update
      Update the own user profile.
      --display-name TEXT
      --given-name TEXT
      --surname TEXT
      --mobile-phone PHONE
      --job-title TEXT
      --department TEXT
      --office-location TEXT

━━━ ONEDRIVE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  onedrive-list [folder]
      List files in a OneDrive folder (default: /)

  onedrive-upload <local_path> <remote_path>
      Upload a local file to OneDrive.

  onedrive-download <remote_path> <local_path>
      Download a file from OneDrive.

  onedrive-delete <remote_path>
      Delete a file or folder on OneDrive.

  onedrive-move <old_path> <new_path>
      Move or rename a file/folder on OneDrive.

  onedrive-share <remote_path>
      Create a share link for a file.
      --anyone        Share with anyone (default: organization)
      --edit          Allow editing (default: view-only)

━━━ ONENOTE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  onenote-create <notebook> <section> <title> <html>
      Create a OneNote page.

  notes-list
      List OneNote notebooks.
      --sections <notebook_id>   List sections in a notebook
      --pages <section_id>       List pages in a section

━━━ TEAMS CHAT ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  chat-list [N]
      List chats (default: 20).

  chat-create
      Create a new chat.
      --members EMAIL[,...]   Other participant(s). One address → oneOnOne chat;
                              two or more → group chat (required)
      --topic TEXT            Chat topic (only for group chats)

  chat-send <chat_id>
      Send a message to a chat.
      --body TEXT   Message text, required

  chat-read <chat_id> [N]
      List N recent messages in a chat (default: 20).

━━━ ONLINE MEETINGS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  meeting-create
      Create an online meeting.
      --subject TEXT   Meeting title, required
      --start ISO      Start datetime (ISO 8601), required
      --end ISO        End datetime (ISO 8601), required
      --participants EMAIL[,...]   Comma-separated attendees

  meeting-list [N]
      List upcoming online meetings (N = days into the future, default: 30).

  meeting-read <meeting_id>
      Show details of an online meeting / calendar event.

  meeting-delete <meeting_id>
      Delete an online meeting / calendar event.

━━━ BOOKINGS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  booking-businesses
      List Microsoft Bookings businesses in the tenant.

  booking-list <business_id> [N]
      List appointments for a Bookings business (default: 50).

  booking-read <business_id> <booking_id>
      Show details of a Bookings appointment.

  booking-create <business_id>
      Create a Bookings appointment.
      --service-id ID        Service ID (required)
      --start ISO            Start datetime (required)
      --end ISO              End datetime (required)
      --customer-name TEXT
      --customer-email EMAIL
      --customer-phone PHONE
      --notes TEXT
      --staff ID[,...]       Staff member IDs (comma-separated)

  booking-cancel <business_id> <booking_id>
      Cancel a Bookings appointment.
      --reason TEXT   Cancellation message

━━━ SHAREPOINT SITES ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  sites-list [N]
      List SharePoint sites accessible to the user (default: 20).

  sites-search <query>
      Search SharePoint sites by keyword.

━━━ EXCEL / WORD / POWERPOINT ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  excel-update <remote_path> <sheet> <range> [value ...]
      Update Excel cells

  word-update <remote_path> key=value [key=value ...]
      Replace placeholders in a Word document

  ppt-update <remote_path> <slide_number> key=value [key=value ...]
      Replace text in a PowerPoint slide (0-indexed)
"""


def _print_json(data):
    print(json.dumps(data, indent=2, default=str))


def _print_tasks(tasks, show_list=False):
    """Pretty-print a list of task dicts."""
    if not tasks:
        print("No tasks found.")
        return
    for t in tasks:
        done = "\u2713" if t.get('is_done') else "\u25cb"
        due = f" (due: {t['due'][:10]})" if t.get('due') else ""
        list_info = f" [{t.get('list_name', '')}]" if show_list else ""
        print(f"{done} {t['title']}{due}{list_info}")
        print(f"  ID: {t['id']}")
        if t.get('note'):
            preview = t['note'].replace('\n', ' ')[:100]
            print(f"  Note: {preview}")
        for step in t.get('steps', []):
            s_done = "\u2713" if step['is_done'] else "\u25cb"
            print(f"    {s_done} {step['title']}  (step ID: {step['id']})")


def main():
    if len(sys.argv) < 2:
        print("OpenClaw M365 CLI ready.")
        print(USAGE)
        sys.exit(0)

    cmd = sys.argv[1]
    args = sys.argv[2:]

    # auth-login does not need an existing session
    if cmd == "auth-login":
        try:
            from m365_openclaw.core import M365Client
            M365Client(force_reauth=True)
            print("Authentication successful. Token cached for headless use.")
        except Exception as exc:
            print(f"Authentication error: {exc}", file=sys.stderr)
            sys.exit(1)
        return

    try:
        from m365_openclaw.core import M365Client
        client = M365Client()
    except EnvironmentError as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        sys.exit(1)

    try:
        # ── Mail ─────────────────────────────────────────────────────────────

        if cmd == "mail-list":
            parser = argparse.ArgumentParser(prog='m365 mail-list', add_help=False)
            parser.add_argument('limit', nargs='?', type=int, default=20)
            parser.add_argument('--sort', default='new-old',
                                choices=['new-old', 'old-new'])
            parser.add_argument('--unread', action='store_true')
            parser.add_argument('--folder', default='inbox')
            parser.add_argument('--search-by-email', default=None)
            parser.add_argument('--search-by-subject', default=None)
            parser.add_argument('--search', default=None)
            parser.add_argument('--count', action='store_true')
            pargs = parser.parse_args(args)
            if pargs.count:
                result = client.list_mail(
                    limit=pargs.limit, folder=pargs.folder,
                    unread_only=pargs.unread, sort=pargs.sort,
                    search_by_email=pargs.search_by_email,
                    search_by_subject=pargs.search_by_subject,
                    search=pargs.search, count_only=True,
                )
                print(f"Count: {result.get('count', 0)}")
            else:
                msgs = client.list_mail(
                    limit=pargs.limit, folder=pargs.folder,
                    unread_only=pargs.unread, sort=pargs.sort,
                    search_by_email=pargs.search_by_email,
                    search_by_subject=pargs.search_by_subject,
                    search=pargs.search,
                )
                if not msgs:
                    print("No messages found.")
                else:
                    for m in msgs:
                        tag = "" if m['is_read'] else "[NEW] "
                        print(f"{tag}{m['date']} | ID: {m['id']} | {m['from']}: {m['subject']}")

        elif cmd == "mail-send":
            parser = argparse.ArgumentParser(prog='m365 mail-send')
            parser.add_argument('--to', required=True,
                                help='Recipient(s), comma-separated')
            parser.add_argument('--subject', required=True)
            parser.add_argument('--body', required=True)
            ct_group = parser.add_mutually_exclusive_group()
            ct_group.add_argument('--html', action='store_true',
                                  help='Send body as HTML (contentType: HTML)')
            ct_group.add_argument('--text', action='store_true',
                                  help='Send body as plain text (contentType: Text, default)')
            parser.add_argument('--cc', default='')
            parser.add_argument('--bcc', default='')
            parser.add_argument('--priority', default='Normal',
                                choices=['High', 'Normal', 'Low'])
            parser.add_argument('--attach', action='append', default=[],
                                dest='attachments', metavar='FILE')
            parser.add_argument('--debug', action='store_true',
                                help='Return full Graph JSON response instead of a status message')
            pargs = parser.parse_args(args)
            to_list = [a.strip() for a in pargs.to.split(',') if a.strip()]
            cc_list = [a.strip() for a in pargs.cc.split(',') if a.strip()]
            bcc_list = [a.strip() for a in pargs.bcc.split(',') if a.strip()]
            content_type = 'HTML' if pargs.html else 'Text'
            result = client.send_mail(
                to_list,
                pargs.subject, pargs.body,
                cc=cc_list or None, bcc=bcc_list or None,
                importance=pargs.priority,
                attachments=pargs.attachments or None,
                content_type=content_type,
                debug=pargs.debug,
            )
            print(json.dumps(result, indent=2) if isinstance(result, dict) else result)

        elif cmd == "mail-read":
            parser = argparse.ArgumentParser(prog='m365 mail-read')
            parser.add_argument('message_id')
            parser.add_argument('--full-body', action='store_true')
            parser.add_argument('--read-header', action='store_true')
            parser.add_argument('--mark-read', action='store_true')
            parser.add_argument('--mark-unread', action='store_true')
            pargs = parser.parse_args(args)
            result = client.read_mail(
                pargs.message_id,
                mark_as_read=pargs.mark_read,
                mark_as_unread=pargs.mark_unread,
                full_body=pargs.full_body,
                headers_only=pargs.read_header,
            )
            _print_json(result)

        elif cmd == "mail-reply":
            parser = argparse.ArgumentParser(prog='m365 mail-reply')
            parser.add_argument('message_id')
            parser.add_argument('--body', required=True)
            parser.add_argument('--reply-all', action='store_true')
            pargs = parser.parse_args(args)
            if pargs.reply_all:
                print(client.reply_all_mail(pargs.message_id, pargs.body))
            else:
                print(client.reply_mail(pargs.message_id, pargs.body))

        elif cmd == "mail-forward":
            parser = argparse.ArgumentParser(prog='m365 mail-forward')
            parser.add_argument('message_id')
            parser.add_argument('--to', required=True)
            parser.add_argument('--body', default='')
            pargs = parser.parse_args(args)
            to_list = [a.strip() for a in pargs.to.split(',') if a.strip()]
            print(client.forward_mail(pargs.message_id, to_list, pargs.body))

        elif cmd == "mail-handle":
            parser = argparse.ArgumentParser(prog='m365 mail-handle')
            parser.add_argument('message_id')
            group = parser.add_mutually_exclusive_group(required=True)
            group.add_argument('--delete', action='store_true')
            group.add_argument('--archive', action='store_true')
            group.add_argument('--move', metavar='FOLDER')
            pargs = parser.parse_args(args)
            if pargs.delete:
                print(client.delete_mail(pargs.message_id))
            elif pargs.archive:
                print(client.move_mail(pargs.message_id, 'archive'))
            else:
                print(client.move_mail(pargs.message_id, pargs.move))

        elif cmd == "mailbox-handle":
            parser = argparse.ArgumentParser(prog='m365 mailbox-handle')
            group = parser.add_mutually_exclusive_group(required=True)
            group.add_argument('--folder-list', action='store_true', default=False)
            group.add_argument('--folder-create', metavar='NAME')
            group.add_argument('--folder-delete', metavar='NAME_OR_ID')
            group.add_argument('--folder-rename', nargs=2, metavar=('ID', 'NAME'))
            pargs = parser.parse_args(args)
            if pargs.folder_list:
                folders = client.list_mail_folders()
                if not folders:
                    print("No mail folders found.")
                else:
                    for f in folders:
                        print(f"{f['id']}\t{f['name']}")
            elif pargs.folder_create:
                print(client.create_mail_folder(pargs.folder_create))
            elif pargs.folder_delete:
                print(client.delete_mail_folder(pargs.folder_delete))
            else:
                fid, new_name = pargs.folder_rename
                print(client.rename_mail_folder(fid, new_name))

        # ── Calendar ─────────────────────────────────────────────────────────

        elif cmd == "calendar-list":
            days = int(args[0]) if args and args[0].lstrip('-').isdigit() else 7
            events = client.get_calendar_events(days=days)
            if not events:
                print(f"No events in the next {days} days.")
            else:
                print(f"Upcoming events (next {days} days):")
                for e in events:
                    loc = f" @ {e['location']}" if e.get('location') else ""
                    print(f"- {e['start']} | {e['subject']}{loc} | {e['id']}")

        elif cmd == "calendar-create":
            parser = argparse.ArgumentParser(prog='m365 calendar-create')
            parser.add_argument('--subject', required=True, help='Event title')
            parser.add_argument('--start', required=True, dest='start_iso',
                                help='Start datetime (ISO 8601)')
            parser.add_argument('--end', required=True, dest='end_iso',
                                help='End datetime (ISO 8601)')
            parser.add_argument('--body', default='', help='Event description')
            ct_group = parser.add_mutually_exclusive_group()
            ct_group.add_argument('--html', action='store_true',
                                  help='Body as HTML (contentType: HTML, default)')
            ct_group.add_argument('--text', action='store_true',
                                  help='Body as plain text (contentType: Text)')
            parser.add_argument('--location', default='')
            parser.add_argument('--required', default='', dest='required_attendees')
            parser.add_argument('--optional', default='', dest='optional_attendees')
            parser.add_argument('--private', action='store_true')
            parser.add_argument('--reminder', type=int, metavar='MINUTES')
            parser.add_argument('--attach', metavar='FILE', default=None)
            pargs = parser.parse_args(args)
            req_att = [e.strip() for e in pargs.required_attendees.split(',')
                       if e.strip()]
            opt_att = [e.strip() for e in pargs.optional_attendees.split(',')
                       if e.strip()]
            body_content_type = 'Text' if pargs.text else 'HTML'
            print(client.create_calendar_event(
                pargs.subject, pargs.start_iso, pargs.end_iso,
                body=pargs.body, location=pargs.location,
                required_attendees=req_att or None,
                optional_attendees=opt_att or None,
                is_private=pargs.private,
                reminder_minutes=pargs.reminder,
                attachment=pargs.attach,
                body_content_type=body_content_type,
            ))

        elif cmd == "calendar-read":
            parser = argparse.ArgumentParser(prog='m365 calendar-read')
            parser.add_argument('event_id')
            parser.add_argument('--participants', action='store_true')
            pargs = parser.parse_args(args)
            event = client.get_calendar_event(pargs.event_id)
            if pargs.participants:
                print(f"Event: {event.get('subject')}")
                print("Attendees:")
                for a in event.get('attendees', []):
                    print(f"  [{a.get('status', '?'):12}] "
                          f"{a.get('name')} <{a.get('email')}> ({a.get('type')})")
            else:
                _print_json(event)

        elif cmd == "calendar-handle":
            parser = argparse.ArgumentParser(prog='m365 calendar-handle')
            parser.add_argument('event_id')
            group = parser.add_mutually_exclusive_group(required=True)
            group.add_argument('--confirm-accept', action='store_true')
            group.add_argument('--confirm-tentative', action='store_true')
            group.add_argument('--confirm-deny', action='store_true')
            group.add_argument('--cancel', action='store_true')
            group.add_argument('--cancel-no-info', action='store_true')
            group.add_argument('--delete', action='store_true')
            pargs = parser.parse_args(args)
            if pargs.confirm_accept:
                print(client.respond_calendar_event(pargs.event_id, 'accept'))
            elif pargs.confirm_tentative:
                print(client.respond_calendar_event(pargs.event_id, 'tentativelyAccept'))
            elif pargs.confirm_deny:
                print(client.respond_calendar_event(pargs.event_id, 'decline'))
            elif pargs.cancel or pargs.cancel_no_info:
                print(client.cancel_calendar_event(pargs.event_id))
            elif pargs.delete:
                print(client.delete_calendar_event(pargs.event_id))

        # ── Contacts ─────────────────────────────────────────────────────────

        elif cmd == "contact-list":
            parser = argparse.ArgumentParser(prog='m365 contact-list', add_help=False)
            parser.add_argument('limit', nargs='?', type=int, default=100)
            parser.add_argument('--sort', default=None,
                                choices=['by-last', 'by-first'])
            parser.add_argument('--count', action='store_true')
            parser.add_argument('--id-only', action='store_true',
                                help='Print only contact IDs, one per line')
            parser.add_argument('--short', action='store_true',
                                help='Print ID, First Name, Last Name and Work Phone only')
            pargs = parser.parse_args(args)
            contacts = client.list_contacts(limit=pargs.limit)
            if pargs.sort == 'by-last':
                contacts.sort(key=lambda c: c.get('last_name', '').lower())
            elif pargs.sort == 'by-first':
                contacts.sort(key=lambda c: c.get('first_name', '').lower())
            if pargs.count:
                print(f"Count: {len(contacts)}")
            elif not contacts:
                print("No contacts found.")
            elif pargs.id_only:
                for c in contacts:
                    print(c['id'])
            elif pargs.short:
                header = (
                    f"{'ID':<48} {'First':<15} {'Last':<15} {'Work Phone':<18}"
                )
                print(header)
                print("-" * len(header))
                for c in contacts:
                    print(
                        f"{c['id']:<48} {c['first_name']:<15} "
                        f"{c['last_name']:<15} {c['work_phone']}"
                    )
            else:
                header = (
                    f"{'ID':<48} {'First':<15} {'Last':<15} {'Work Email':<30} "
                    f"{'Home Email':<25} {'Work Phone':<18} {'Mobile'}"
                )
                print(header)
                print("-" * len(header))
                for c in contacts:
                    print(
                        f"{c['id']:<48} {c['first_name']:<15} {c['last_name']:<15} "
                        f"{c['work_email']:<30} {c['home_email']:<25} "
                        f"{c['work_phone']:<18} {c['mobile_phone']}"
                    )

        elif cmd == "contact-read":
            if not args:
                print("Usage: m365 contact-read <contact_id>")
                sys.exit(1)
            contact = client.get_contact_by_id(args[0])
            if not contact:
                print(f"No contact found with ID '{args[0]}'.")
            else:
                _print_json(contact)

        elif cmd == "contact-edit":
            parser = argparse.ArgumentParser(prog='m365 contact-edit')
            parser.add_argument('contact_id')
            parser.add_argument('--given-name', default=None)
            parser.add_argument('--surname', default=None)
            parser.add_argument('--email-business', default=None)
            parser.add_argument('--email-personal', default=None)
            parser.add_argument('--phone-mobile', default=None)
            parser.add_argument('--phone-business', default=None)
            parser.add_argument('--phone-home', default=None)
            parser.add_argument('--birthday', default=None)
            parser.add_argument('--notes', default=None)
            parser.add_argument('--company', default=None)
            parser.add_argument('--job-title', default=None)
            pargs = parser.parse_args(args)
            print(client.update_contact(
                pargs.contact_id,
                given_name=pargs.given_name,
                surname=pargs.surname,
                email_business=pargs.email_business,
                email_personal=pargs.email_personal,
                phone_mobile=pargs.phone_mobile,
                phone_business=pargs.phone_business,
                phone_home=pargs.phone_home,
                birthday=pargs.birthday,
                notes=pargs.notes,
                company=pargs.company,
                job_title=pargs.job_title,
            ))

        elif cmd == "contact-create":
            parser = argparse.ArgumentParser(prog='m365 contact-create')
            parser.add_argument('given_name', help='First name')
            parser.add_argument('surname', help='Last name')
            parser.add_argument('--email-business', default=None)
            parser.add_argument('--email-personal', default=None)
            parser.add_argument('--phone-mobile', default=None)
            parser.add_argument('--phone-business', default=None)
            parser.add_argument('--phone-home', default=None)
            parser.add_argument('--birthday', default=None,
                                metavar='YYYY-MM-DDT00:00:00Z')
            parser.add_argument('--notes', default=None)
            parser.add_argument('--company', default=None)
            parser.add_argument('--job-title', default=None)
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
            parser.add_argument('--anniversary', default=None,
                                metavar='YYYY-MM-DDT00:00:00Z')
            parser.add_argument('--website', default=None)
            parser.add_argument('--work-website', default=None)
            parser.add_argument('--spouse', default=None)
            pargs = parser.parse_args(args)
            print(client.create_contact(
                pargs.given_name, pargs.surname,
                work_email=pargs.email_business,
                home_email=pargs.email_personal,
                work_phone=pargs.phone_business,
                home_phone=pargs.phone_home,
                mobile_phone=pargs.phone_mobile,
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

        elif cmd == "contact-photo":
            parser = argparse.ArgumentParser(prog='m365 contact-photo')
            parser.add_argument('contact_id')
            group = parser.add_mutually_exclusive_group(required=True)
            group.add_argument('--upload', metavar='FILE')
            group.add_argument('--delete', action='store_true')
            group.add_argument('--download', metavar='SAVE_PATH')
            pargs = parser.parse_args(args)
            if pargs.upload:
                print(client.set_contact_photo(pargs.contact_id, pargs.upload))
            elif pargs.delete:
                print(client.delete_contact_photo(pargs.contact_id))
            else:
                print(client.get_contact_photo(pargs.contact_id, pargs.download))

        elif cmd == "contactlist-list":
            folders = client.list_contact_folders()
            if not folders:
                print("No contact folders found.")
            else:
                for f in folders:
                    print(f"{f['id']:<50} {f['name']}")

        elif cmd == "contactlist-create":
            if not args:
                print("Usage: m365 contactlist-create <name>")
                sys.exit(1)
            print(client.create_contact_folder(args[0]))

        elif cmd == "contactlist-delete":
            if not args:
                print("Usage: m365 contactlist-delete <folder_id>")
                sys.exit(1)
            print(client.delete_contact_folder(args[0]))

        elif cmd == "contact-delete":
            if not args:
                print("Usage: m365 contact-delete <contact_id>")
                sys.exit(1)
            print(client.delete_contact(args[0]))

        # ── Tasks ─────────────────────────────────────────────────────────────

        elif cmd == "task-list":
            parser = argparse.ArgumentParser(prog='m365 task-list', add_help=False)
            parser.add_argument('limit', nargs='?', type=int, default=50)
            parser.add_argument('--list-id', default=None)
            parser.add_argument('--status', default=None,
                                choices=['notStarted', 'inProgress', 'completed'])
            parser.add_argument('--count', action='store_true')
            pargs = parser.parse_args(args)
            if pargs.list_id:
                tasks = client.todo_list_tasks(pargs.list_id)
            else:
                tasks = client.todo_get_all_tasks()
            if pargs.status:
                tasks = [t for t in tasks if t.get('status') == pargs.status]
            if pargs.count:
                print(f"Count: {len(tasks)}")
            else:
                _print_tasks(tasks, show_list=(pargs.list_id is None))

        elif cmd == "task-create":
            parser = argparse.ArgumentParser(prog='m365 task-create')
            parser.add_argument('--title', required=True)
            parser.add_argument('--list-id', default=None)
            parser.add_argument('--due', default=None, metavar='YYYY-MM-DDTHH:MM:SS')
            parser.add_argument('--body', default=None)
            parser.add_argument('--priority', default=None,
                                choices=['low', 'normal', 'high'])
            pargs = parser.parse_args(args)
            list_id = pargs.list_id or client.todo_get_default_list_id()
            if not list_id:
                print("No task list found. Create one first with: m365 tasklist-create <name>",
                      file=sys.stderr)
                sys.exit(1)
            print(client.todo_create_task(
                list_id, pargs.title,
                note=pargs.body,
                due_date=pargs.due,
            ))

        elif cmd == "task-read":
            parser = argparse.ArgumentParser(prog='m365 task-read')
            parser.add_argument('task_id')
            parser.add_argument('--list-id', default=None)
            pargs = parser.parse_args(args)
            list_id = pargs.list_id or client.todo_get_default_list_id()
            if not list_id:
                print("No task list found.", file=sys.stderr)
                sys.exit(1)
            task = client.todo_get_task(list_id, pargs.task_id)
            if task:
                _print_json(task)
            else:
                print(f"Task '{pargs.task_id}' not found.")

        elif cmd == "task-edit":
            parser = argparse.ArgumentParser(prog=f'm365 {cmd}')
            parser.add_argument('task_id')
            parser.add_argument('--list-id', default=None)
            parser.add_argument('--title', default=None)
            parser.add_argument('--status', default=None,
                                choices=['notStarted', 'inProgress', 'completed'])
            parser.add_argument('--due', default=None, metavar='YYYY-MM-DDTHH:MM:SS')
            parser.add_argument('--complete', action='store_true')
            checklist_group = parser.add_mutually_exclusive_group()
            checklist_group.add_argument('--checklist-add', default=None, metavar='TEXT',
                                         help='Add a checklist item with the given title')
            checklist_group.add_argument('--checklist-complete', default=None, metavar='ID',
                                         help='Mark a checklist item as completed')
            checklist_group.add_argument('--checklist-delete', default=None, metavar='ID',
                                         help='Delete a checklist item')
            pargs = parser.parse_args(args)
            list_id = pargs.list_id or client.todo_get_default_list_id()
            if not list_id:
                print("No task list found.", file=sys.stderr)
                sys.exit(1)
            if pargs.checklist_add:
                print(client.todo_add_step(list_id, pargs.task_id, pargs.checklist_add))
            elif pargs.checklist_complete:
                print(client.todo_complete_step(list_id, pargs.task_id, pargs.checklist_complete))
            elif pargs.checklist_delete:
                print(client.todo_delete_checklist_item(list_id, pargs.task_id, pargs.checklist_delete))
            elif pargs.complete:
                print(client.todo_complete_task(list_id, pargs.task_id))
            else:
                print(client.todo_update_task(
                    list_id, pargs.task_id,
                    title=pargs.title,
                    due_date=pargs.due,
                ))

        elif cmd == "task-delete":
            parser = argparse.ArgumentParser(prog='m365 task-delete')
            parser.add_argument('task_id')
            parser.add_argument('--list-id', default=None)
            pargs = parser.parse_args(args)
            list_id = pargs.list_id or client.todo_get_default_list_id()
            if not list_id:
                print("No task list found.", file=sys.stderr)
                sys.exit(1)
            print(client.todo_delete_task(list_id, pargs.task_id))

        elif cmd == "tasklist-list":
            lists = client.todo_list_task_lists()
            if not lists:
                print("No task lists found.")
            else:
                print(f"{'ID':<50} Name")
                print("-" * 70)
                for lst in lists:
                    shared = " [shared]" if lst.get('is_shared') else ""
                    print(f"{lst['id']:<50} {lst['name']}{shared}")

        elif cmd == "tasklist-create":
            if not args:
                print("Usage: m365 tasklist-create <name>")
                sys.exit(1)
            print(client.todo_create_task_list(args[0]))

        # ── User ─────────────────────────────────────────────────────────────

        elif cmd == "user-read":
            _print_json(client.get_user())

        elif cmd == "user-update":
            parser = argparse.ArgumentParser(prog='m365 user-update')
            parser.add_argument('--display-name', default=None)
            parser.add_argument('--given-name', default=None)
            parser.add_argument('--surname', default=None)
            parser.add_argument('--mobile-phone', default=None)
            parser.add_argument('--job-title', default=None)
            parser.add_argument('--department', default=None)
            parser.add_argument('--office-location', default=None)
            pargs = parser.parse_args(args)
            print(client.update_user(
                display_name=pargs.display_name,
                given_name=pargs.given_name,
                surname=pargs.surname,
                mobile_phone=pargs.mobile_phone,
                job_title=pargs.job_title,
                department=pargs.department,
                office_location=pargs.office_location,
            ))

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

        elif cmd == "onedrive-upload":
            if len(args) < 2:
                print("Usage: m365 onedrive-upload <local_path> <remote_path>")
                sys.exit(1)
            print(client.onedrive_upload(args[0], args[1]))

        elif cmd == "onedrive-download":
            if len(args) < 2:
                print("Usage: m365 onedrive-download <remote_path> <local_path>")
                sys.exit(1)
            print(client.onedrive_download(args[0], args[1]))

        elif cmd == "onedrive-delete":
            if not args:
                print("Usage: m365 onedrive-delete <remote_path>")
                sys.exit(1)
            print(client.onedrive_delete(args[0]))

        elif cmd == "onedrive-move":
            if len(args) < 2:
                print("Usage: m365 onedrive-move <old_path> <new_path>")
                sys.exit(1)
            print(client.onedrive_move(args[0], args[1]))

        elif cmd == "onedrive-share":
            parser = argparse.ArgumentParser(prog='m365 onedrive-share')
            parser.add_argument('remote_path')
            parser.add_argument('--anyone', action='store_true')
            parser.add_argument('--edit', action='store_true')
            pargs = parser.parse_args(args)
            print(client.onedrive_share(pargs.remote_path,
                                        anyone=pargs.anyone, edit=pargs.edit))

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

        elif cmd == "task-handle":
            parser = argparse.ArgumentParser(prog='m365 task-handle')
            parser.add_argument('task_id')
            parser.add_argument('--list-id', default=None)
            parser.add_argument('--to', required=True, metavar='TO_LIST_ID',
                                help='Destination list ID to move the task to')
            pargs = parser.parse_args(args)
            list_id = pargs.list_id or client.todo_get_default_list_id()
            if not list_id:
                print("No task list found.", file=sys.stderr)
                sys.exit(1)
            print(client.todo_move_task(list_id, pargs.task_id, pargs.to))

        # ── OneNote extended listing ──────────────────────────────────────────

        elif cmd == "notes-list":
            parser = argparse.ArgumentParser(prog='m365 notes-list', add_help=False)
            parser.add_argument('--sections', metavar='NOTEBOOK_ID', default=None)
            parser.add_argument('--pages', metavar='SECTION_ID', default=None)
            pargs = parser.parse_args(args)
            if pargs.pages:
                pages = client.notes_list_pages(pargs.pages)
                if not pages:
                    print("No pages found.")
                else:
                    for p in pages:
                        print(f"{p['created'][:10]}  {p['title']}")
                        print(f"  ID: {p['id']}")
            elif pargs.sections:
                sections = client.notes_list_sections(pargs.sections)
                if not sections:
                    print("No sections found.")
                else:
                    for s in sections:
                        print(f"{s['name']:<40} ID: {s['id']}")
            else:
                notebooks = client.notes_list_notebooks()
                if not notebooks:
                    print("No notebooks found.")
                else:
                    for nb in notebooks:
                        print(f"{nb['name']:<40} ID: {nb['id']}")

        # ── Teams Chat ────────────────────────────────────────────────────────

        elif cmd == "chat-list":
            limit = int(args[0]) if args and args[0].isdigit() else 20
            chats = client.chat_list(limit=limit)
            if not chats:
                print("No chats found.")
            else:
                for c in chats:
                    members = ", ".join(c['members'])
                    topic = f" [{c['topic']}]" if c.get('topic') else ""
                    print(f"[{c['type']}]{topic} ID: {c['id']}")
                    print(f"  Members: {members}")

        elif cmd == "chat-create":
            parser = argparse.ArgumentParser(prog='m365 chat-create')
            parser.add_argument('--members', required=True,
                                help='Comma-separated list of email addresses or user IDs')
            parser.add_argument('--topic', default=None)
            pargs = parser.parse_args(args)
            members = [m.strip() for m in pargs.members.split(',') if m.strip()]
            if len(members) < 1:
                print("At least 1 member is required to create a chat.", file=sys.stderr)
                sys.exit(1)
            print(client.chat_create(members, topic=pargs.topic))

        elif cmd == "chat-send":
            parser = argparse.ArgumentParser(prog='m365 chat-send')
            parser.add_argument('chat_id')
            parser.add_argument('--body', required=True)
            pargs = parser.parse_args(args)
            print(client.chat_send(pargs.chat_id, pargs.body))

        elif cmd == "chat-read":
            parser = argparse.ArgumentParser(prog='m365 chat-read', add_help=False)
            parser.add_argument('chat_id')
            parser.add_argument('limit', nargs='?', type=int, default=20)
            pargs = parser.parse_args(args)
            messages = client.chat_read(pargs.chat_id, limit=pargs.limit)
            if not messages:
                print("No messages found.")
            else:
                for m in messages:
                    print(f"{m['created'][:19]}  {m['sender']}: {m['body'][:120]}")

        # ── Online Meetings ───────────────────────────────────────────────────

        elif cmd == "meeting-create":
            parser = argparse.ArgumentParser(prog='m365 meeting-create')
            parser.add_argument('--subject', required=True)
            parser.add_argument('--start', required=True, dest='start_iso')
            parser.add_argument('--end', required=True, dest='end_iso')
            parser.add_argument('--participants', default='')
            pargs = parser.parse_args(args)
            participants = [p.strip() for p in pargs.participants.split(',')
                            if p.strip()]
            result = client.meeting_create(
                pargs.subject, pargs.start_iso, pargs.end_iso,
                participants=participants or None,
            )
            print(f"Meeting '{result['subject']}' created")
            print(f"  ID:       {result['id']}")
            print(f"  Join URL: {result['join_url']}")
            if result.get('join_id'):
                print(f"  Join ID:  {result['join_id']}")

        elif cmd == "meeting-list":
            days = int(args[0]) if args and args[0].isdigit() else 30
            meetings = client.meeting_list(days=days)
            if not meetings:
                print(f"No online meetings in the next {days} days.")
            else:
                for m in meetings:
                    print(f"{m['start'][:16]}  {m['subject']}")
                    print(f"  ID: {m['id']}")
                    if m.get('join_url'):
                        print(f"  Join: {m['join_url']}")

        elif cmd == "meeting-read":
            if not args:
                print("Usage: m365 meeting-read <meeting_id>")
                sys.exit(1)
            _print_json(client.meeting_read(args[0]))

        elif cmd == "meeting-delete":
            if not args:
                print("Usage: m365 meeting-delete <meeting_id>")
                sys.exit(1)
            print(client.meeting_delete(args[0]))

        # ── Microsoft Bookings ────────────────────────────────────────────────

        elif cmd == "booking-businesses":
            businesses = client.booking_businesses()
            if not businesses:
                print("No Bookings businesses found.")
            else:
                for b in businesses:
                    print(f"{b['name']:<40} ID: {b['id']}")
                    if b.get('email'):
                        print(f"  Email: {b['email']}")

        elif cmd == "booking-list":
            if not args:
                print("Usage: m365 booking-list <business_id> [N]")
                sys.exit(1)
            business_id = args[0]
            limit = int(args[1]) if len(args) > 1 and args[1].isdigit() else 50
            bookings = client.booking_list(business_id, limit=limit)
            if not bookings:
                print("No bookings found.")
            else:
                for b in bookings:
                    print(f"{b['start'][:16]}  {b['service_name']}  – {b['customer']}")
                    print(f"  ID: {b['id']}")

        elif cmd == "booking-read":
            if len(args) < 2:
                print("Usage: m365 booking-read <business_id> <booking_id>")
                sys.exit(1)
            _print_json(client.booking_read(args[0], args[1]))

        elif cmd == "booking-create":
            parser = argparse.ArgumentParser(prog='m365 booking-create')
            parser.add_argument('business_id')
            parser.add_argument('--service-id', required=True)
            parser.add_argument('--start', required=True, dest='start_iso')
            parser.add_argument('--end', required=True, dest='end_iso')
            parser.add_argument('--customer-name', default='')
            parser.add_argument('--customer-email', default='')
            parser.add_argument('--customer-phone', default='')
            parser.add_argument('--notes', default='')
            parser.add_argument('--staff', default='',
                                help='Comma-separated staff member IDs')
            pargs = parser.parse_args(args)
            staff_ids = [s.strip() for s in pargs.staff.split(',') if s.strip()]
            print(client.booking_create(
                pargs.business_id, pargs.service_id,
                pargs.start_iso, pargs.end_iso,
                customer_name=pargs.customer_name,
                customer_email=pargs.customer_email,
                customer_phone=pargs.customer_phone,
                notes=pargs.notes,
                staff_ids=staff_ids or None,
            ))

        elif cmd == "booking-cancel":
            parser = argparse.ArgumentParser(prog='m365 booking-cancel')
            parser.add_argument('business_id')
            parser.add_argument('booking_id')
            parser.add_argument('--reason', default='')
            pargs = parser.parse_args(args)
            print(client.booking_cancel(pargs.business_id, pargs.booking_id,
                                        reason=pargs.reason))

        # ── SharePoint Sites ──────────────────────────────────────────────────

        elif cmd == "sites-list":
            limit = int(args[0]) if args and args[0].isdigit() else 20
            sites = client.sites_list(limit=limit)
            if not sites:
                print("No sites found.")
            else:
                for s in sites:
                    print(f"{s['name']:<40} {s['url']}")
                    print(f"  ID: {s['id']}")

        elif cmd == "sites-search":
            if not args:
                print("Usage: m365 sites-search <query>")
                sys.exit(1)
            sites = client.sites_search(args[0])
            if not sites:
                print("No sites found.")
            else:
                for s in sites:
                    print(f"{s['name']:<40} {s['url']}")
                    print(f"  ID: {s['id']}")

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
