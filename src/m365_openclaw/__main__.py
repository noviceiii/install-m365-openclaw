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
