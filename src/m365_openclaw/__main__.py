"""
__main__.py – CLI entry point for the OpenClaw M365 skill (v0.3.0).

Invoked as:  python -m m365_openclaw <command> [arguments...]
Or via the  m365  wrapper script placed in ~/.local/bin.

All v0.1.0 commands remain 100% compatible.
"""

import argparse
import json
import sys

USAGE = """\
OpenClaw M365 CLI v0.3.0 – Microsoft 365 for agents

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
      Update Excel cells

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
