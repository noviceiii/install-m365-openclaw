# SKILL: m365-graph

OpenClaw skill for unattended Microsoft 365 access via Graph API.

**Version:** 0.5.0 (Hybrid: Delegated Device-Code + Application Client-Credentials)

## Executable
m365

## Description
Enables OpenClaw agents to send/read mail, manage calendar events, manage contacts,
access OneDrive files, create OneNote pages, edit Excel/Word/PowerPoint documents,
manage Microsoft To Do tasks, create Teams chats, schedule online meetings, manage
Bookings appointments, and list SharePoint sites.

Two authentication flows are used:
- **Delegated (device-code)** – for all features except mail sending. The user signs
  in once interactively; the app then works fully headless using cached refresh tokens
  via `GET /me/…`.
- **Application (client-credentials)** – used exclusively for `mail-send`. Exchange
  Online requires `POST /users/{UPN}/sendMail` with an application token to reliably
  send mail from unattended service applications. This flow requires `CLIENT_SECRET`
  and `MAIL_SENDER_UPN` in `.env`.

## Configuration
Credentials stored in: ~/.openclaw/skills/m365-graph/.env

Required environment variables (all features):
- TENANT_ID
- CLIENT_ID
- TOKEN_CACHE_PATH

Additional variables required for `mail-send` only:
- CLIENT_SECRET   – app client secret (application/client-credentials flow)
- MAIL_SENDER_UPN – UPN or e-mail of the mailbox to send from (e.g. sender@example.com)

## Required Microsoft Graph Permissions

### Delegated permissions (device-code flow – all features except mail sending)

All permissions below must be added as **delegated** permissions in Entra ID and
granted admin consent:

`User.Read`, `openid`, `profile`, `offline_access`,
`Mail.ReadWrite`, `Mail.Send`,
`Calendars.ReadWrite`, `Contacts.ReadWrite`, `MailboxFolder.ReadWrite`,
`Tasks.ReadWrite`, `Files.ReadWrite`, `Notes.ReadWrite`,
`Sites.ReadWrite.All`,
`Bookings.Manage.All`, `Bookings.ReadWrite.All`, `BookingsAppointment.ReadWrite.All`,
`Chat.Create`, `Chat.ReadWrite`, `OnlineMeetings.ReadWrite`

### Application permissions (client-credentials flow – mail sending only)

Add the following as an **application** permission in Entra ID and grant admin consent:

`Mail.Send`

> **Note:** The `sensitivity` field is intentionally excluded from `mail-send`
> payloads. Microsoft Graph v1.0 does not support `sensitivity` on
> `microsoft.graph.message` for the `sendMail` action, and Exchange Online returns
> HTTP 400 when it is present.

## First-Time Setup

```bash
m365 auth-login
```

Open `https://microsoft.com/devicelogin` on any device and enter the displayed code.
After sign-in the app runs headless; tokens refresh automatically.

## Commands

### Auth

| Command | Description | Example |
|---------|-------------|---------|
| `auth-login` | Force device-code login / re-authentication | `m365 auth-login` |

### Mail

| Command | Description | Example |
|---------|-------------|---------|
| `mail-list [N]` | List N recent messages (default 20, max 500) | `m365 mail-list 50 --unread --folder Sent` |
| `mail-list --sort new-old\|old-new` | Sort order | `m365 mail-list --sort old-new` |
| `mail-list --unread` | Only unread messages | `m365 mail-list --unread` |
| `mail-list --folder NAME` | List messages from a specific folder | `m365 mail-list --folder Drafts` |
| `mail-list --search-by-email ADDR` | Filter by sender address | `m365 mail-list --search-by-email boss@company.com` |
| `mail-list --search-by-subject TEXT` | Filter by subject | `m365 mail-list --search-by-subject Invoice` |
| `mail-list --search TEXT` | Full-text search | `m365 mail-list --search "project update"` |
| `mail-list --count` | Count messages only | `m365 mail-list --unread --count` |
| `mail-send --to ADDR --subject TEXT --body TEXT` | Send an e-mail (plain text, default) | `m365 mail-send --to boss@co.com --subject "Update" --body "Done"` |
| `mail-send ... --html` | Send body as HTML | `m365 mail-send --to boss@co.com --subject "Hi" --body "<b>Hello</b>" --html` |
| `mail-send ... --text` | Send body as plain text (explicit) | `m365 mail-send --to boss@co.com --subject "Hi" --body "Hello" --text` |
| `mail-send --to ADDR,ADDR --subject TEXT --body TEXT` | Send to multiple recipients | `m365 mail-send --to a@co.com,b@co.com --subject "Update" --body "Done"` |
| `mail-send ... --cc ADDR` | Add CC | `m365 mail-send --to a@b.com --subject S --body B --cc c@b.com` |
| `mail-send ... --bcc ADDR` | Add BCC | `m365 mail-send --to a@b.com --subject S --body B --bcc d@b.com` |
| `mail-send ... --attach FILE` | Attach a file | `m365 mail-send --to a@b.com --subject S --body B --attach /path/f.pdf` |
| `mail-send ... --priority High\|Normal\|Low` | Set priority | `m365 mail-send --to a@b.com --subject S --body B --priority High` |
| `mail-read \<id\>` | Show message details (body max 200 chars) | `m365 mail-read AAMk...` |
| `mail-read \<id\> --full-body` | Show complete message body | `m365 mail-read AAMk... --full-body` |
| `mail-read \<id\> --read-header` | Show headers only | `m365 mail-read AAMk... --read-header` |
| `mail-read \<id\> --mark-read` | Mark message as read | `m365 mail-read AAMk... --mark-read` |
| `mail-read \<id\> --mark-unread` | Mark message as unread | `m365 mail-read AAMk... --mark-unread` |
| `mail-reply \<id\> --body TEXT` | Reply to sender | `m365 mail-reply AAMk... --body "Got it!"` |
| `mail-reply \<id\> --body TEXT --reply-all` | Reply to all | `m365 mail-reply AAMk... --body "Noted" --reply-all` |
| `mail-forward \<id\> --to ADDR` | Forward message | `m365 mail-forward AAMk... --to colleague@co.com` |
| `mail-forward \<id\> --to ADDR --body TEXT` | Forward with comment | `m365 mail-forward AAMk... --to c@co.com --body "FYI"` |
| `mail-handle \<id\> --delete` | Delete a message | `m365 mail-handle AAMk... --delete` |
| `mail-handle \<id\> --archive` | Archive a message | `m365 mail-handle AAMk... --archive` |
| `mail-handle \<id\> --move FOLDER` | Move to folder | `m365 mail-handle AAMk... --move ProjektX` |
| `mailbox-handle --folder-create NAME` | Create mail folder | `m365 mailbox-handle --folder-create ProjektX` |
| `mailbox-handle --folder-delete NAME` | Delete mail folder | `m365 mailbox-handle --folder-delete ProjektX` |
| `mailbox-handle --folder-rename ID NAME` | Rename mail folder | `m365 mailbox-handle --folder-rename AAMk... NewName` |

### Calendar

| Command | Description | Example |
|---------|-------------|---------|
| `calendar-list [N]` | List upcoming events (N = days, default 7); output includes date/time, subject, location, and event ID | `m365 calendar-list 14` |
| `calendar-create --subject TEXT --start ISO --end ISO` | Create event | `m365 calendar-create --subject Meeting --start 2026-04-10T10:00 --end 2026-04-10T11:00` |
| `calendar-create ... --body TEXT` | Add description (HTML by default) | `m365 calendar-create --subject S --start T1 --end T2 --body "Q1 review"` |
| `calendar-create ... --body TEXT --text` | Add description as plain text | `m365 calendar-create --subject S --start T1 --end T2 --body "Q1 review" --text` |
| `calendar-create ... --body TEXT --html` | Add description as HTML (explicit) | `m365 calendar-create --subject S --start T1 --end T2 --body "<b>Q1</b>" --html` |
| `calendar-create ... --location TEXT` | Set location | `m365 calendar-create --subject S --start T1 --end T2 --location "Room 101"` |
| `calendar-create ... --required EMAIL` | Add required attendee | `m365 calendar-create --subject S --start T1 --end T2 --required alice@co.com` |
| `calendar-create ... --optional EMAIL` | Add optional attendee | `m365 calendar-create --subject S --start T1 --end T2 --optional bob@co.com` |
| `calendar-create ... --private` | Mark as private | `m365 calendar-create --subject S --start T1 --end T2 --private` |
| `calendar-create ... --reminder N` | Set reminder (minutes) | `m365 calendar-create --subject S --start T1 --end T2 --reminder 30` |
| `calendar-create ... --attach FILE` | Attach file | `m365 calendar-create --subject S --start T1 --end T2 --attach agenda.pdf` |
| `calendar-read \<id\>` | Show event details | `m365 calendar-read AAES...` |
| `calendar-read \<id\> --participants` | List attendees and status | `m365 calendar-read AAES... --participants` |
| `calendar-handle \<id\> --confirm-accept` | Accept invitation | `m365 calendar-handle AAES... --confirm-accept` |
| `calendar-handle \<id\> --confirm-tentative` | Tentatively accept | `m365 calendar-handle AAES... --confirm-tentative` |
| `calendar-handle \<id\> --confirm-deny` | Decline invitation | `m365 calendar-handle AAES... --confirm-deny` |
| `calendar-handle \<id\> --cancel` | Cancel event | `m365 calendar-handle AAES... --cancel` |
| `calendar-handle \<id\> --delete` | Delete event | `m365 calendar-handle AAES... --delete` |

### Contacts

| Command | Description | Example |
|---------|-------------|---------|
| `contact-list [N]` | List N contacts with ID as first column (default 100) | `m365 contact-list` |
| `contact-list --sort by-last\|by-first` | Sort contacts | `m365 contact-list --sort by-last` |
| `contact-list --count` | Count contacts | `m365 contact-list --count` |
| `contact-list --id-only` | Print only contact IDs, one per line (for scripting) | `m365 contact-list --id-only` |
| `contact-list --short` | Print ID, First Name, Last Name and Work Phone only | `m365 contact-list --short` |
| `contact-read \<id\>` | Show all fields of a contact | `m365 contact-read AAMk...` |
| `contact-edit \<id\> [--field VALUE ...]` | Update contact fields | `m365 contact-edit AAMk... --email-business neu@domain.ch` |
| `contact-create \<first\> \<last\> [options]` | Create new contact | `m365 contact-create Jane Doe --email-business j@work.com` |
| `contact-photo \<id\> --upload FILE` | Upload contact photo | `m365 contact-photo AAMk... --upload foto.jpg` |
| `contact-photo \<id\> --delete` | Delete contact photo | `m365 contact-photo AAMk... --delete` |
| `contact-photo \<id\> --download PATH` | Download contact photo | `m365 contact-photo AAMk... --download /tmp/foto.jpg` |
| `contactlist-list` | List contact folders | `m365 contactlist-list` |
| `contactlist-create \<name\>` | Create contact folder | `m365 contactlist-create Kunden2026` |
| `contactlist-delete \<id\>` | Delete contact folder | `m365 contactlist-delete AAMk...` |
| `contact-delete \<id\>` | Delete a contact by ID | `m365 contact-delete AAMk...` |

### Tasks

| Command | Description | Example |
|---------|-------------|---------|
| `task-list [N]` | List tasks (default 50) | `m365 task-list 50` |
| `task-list --list-id ID` | List tasks in specific list | `m365 task-list --list-id AAMk...` |
| `task-list --status STATUS` | Filter by status | `m365 task-list --status notStarted` |
| `task-list --count` | Count tasks | `m365 task-list --count` |
| `task-create --title TEXT` | Create a task | `m365 task-create --title "Bericht fertig" --due 2026-04-15` |
| `task-create ... --list-id ID` | Create in specific list | `m365 task-create --title T --list-id AAMk...` |
| `task-create ... --due DATETIME` | Set due date | `m365 task-create --title T --due 2026-04-15T00:00:00` |
| `task-create ... --body TEXT` | Add notes | `m365 task-create --title T --body "Details here"` |
| `task-read \<id\>` | Show task details | `m365 task-read ABC123` |
| `task-read \<id\> --list-id ID` | Read from specific list | `m365 task-read ABC123 --list-id AAMk...` |
| `task-update \<id\> [options]` | Update a task | `m365 task-update ABC123 --status completed` |
| `task-update \<id\> --complete` | Mark task as completed | `m365 task-update ABC123 --complete` |
| `task-delete \<id\>` | Delete a task | `m365 task-delete ABC123` |
| `tasklist-list` | List all task lists | `m365 tasklist-list` |
| `tasklist-create \<name\>` | Create a task list | `m365 tasklist-create Privat` |

### User

| Command | Description | Example |
|---------|-------------|---------|
| `user-read` | Show own user profile | `m365 user-read` |
| `user-update [--field VALUE ...]` | Update own user profile | `m365 user-update --mobile-phone +41791234567` |

### OneDrive

| Command | Description | Example |
|---------|-------------|---------|
| `onedrive-list [folder]` | List OneDrive folder | `m365 onedrive-list /Dokumente` |
| `onedrive-upload \<local\> \<remote\>` | Upload to OneDrive | `m365 onedrive-upload ./report.pdf /Finance/report.pdf` |
| `onedrive-download \<remote\> \<local\>` | Download from OneDrive | `m365 onedrive-download /Finance/report.pdf ./local.pdf` |
| `onedrive-delete \<remote\>` | Delete file or folder | `m365 onedrive-delete /alt/bericht.pdf` |
| `onedrive-move \<old\> \<new\>` | Move or rename file/folder | `m365 onedrive-move /alt.pdf /neu/bericht.pdf` |
| `onedrive-share \<remote\>` | Create share link (org) | `m365 onedrive-share /dokument.docx` |
| `onedrive-share \<remote\> --anyone` | Share with anyone | `m365 onedrive-share /dokument.docx --anyone` |
| `onedrive-share \<remote\> --edit` | Allow editing | `m365 onedrive-share /dokument.docx --anyone --edit` |

### Office Documents

| Command | Description | Example |
|---------|-------------|---------|
| `excel-update \<path\> \<sheet\> \<range\> [val ...]` | Update Excel cells | `m365 excel-update Report.xlsx Sheet1 A1:B2 Jan Feb 100 200` |
| `word-update \<path\> key=val ...` | Replace placeholders in Word | `m365 word-update Offer.docx "{{Name}}=Alice"` |
| `ppt-update \<path\> \<slide\> key=val ...` | Replace text in PowerPoint | `m365 ppt-update Deck.pptx 0 "{{Title}}=2026 Strategy"` |

### OneNote

| Command | Description | Example |
|---------|-------------|---------|
| `onenote-create \<nb\> \<sec\> \<title\> \<html\>` | Create OneNote page | `m365 onenote-create WorkNotes April "Sync" "<h1>Notes</h1>"` |
| `notes-list` | List notebooks | `m365 notes-list` |
| `notes-list --sections \<nb_id\>` | List sections in notebook | `m365 notes-list --sections AAMk...` |
| `notes-list --pages \<sec_id\>` | List pages in section | `m365 notes-list --pages AAMk...` |

### Teams Chat

| Command | Description | Example |
|---------|-------------|---------|
| `chat-list [N]` | List chats (default 20) | `m365 chat-list 10` |
| `chat-create --members EMAIL[,...]` | Create a chat (1 address = oneOnOne, 2+ = group) | `m365 chat-create --members alice@co.com` |
| `chat-send \<chat_id\> --body TEXT` | Send a message | `m365 chat-send 19:abc...@thread.v2 --body "Hello team!"` |
| `chat-read \<chat_id\> [N]` | Read messages (default 20) | `m365 chat-read 19:abc...@thread.v2 10` |

### Online Meetings

| Command | Description | Example |
|---------|-------------|---------|
| `meeting-create --subject TEXT --start ISO --end ISO` | Create online meeting | `m365 meeting-create --subject "Standup" --start 2026-04-10T09:00 --end 2026-04-10T09:30` |
| `meeting-create ... --participants EMAIL[,...]` | Add participants | `m365 meeting-create --subject S --start T1 --end T2 --participants alice@co.com` |
| `meeting-list [N]` | List upcoming meetings (N = days, default 30) | `m365 meeting-list 14` |
| `meeting-read \<id\>` | Show meeting details | `m365 meeting-read AAES...` |
| `meeting-delete \<id\>` | Delete meeting | `m365 meeting-delete AAES...` |

### Bookings

| Command | Description | Example |
|---------|-------------|---------|
| `booking-businesses` | List Bookings businesses | `m365 booking-businesses` |
| `booking-list \<business_id\> [N]` | List appointments (default 50) | `m365 booking-list BIZ_ID 20` |
| `booking-read \<business_id\> \<booking_id\>` | Show appointment details | `m365 booking-read BIZ_ID APT_ID` |
| `booking-create \<business_id\> --service-id ID --start ISO --end ISO` | Create appointment | `m365 booking-create BIZ_ID --service-id SVC_ID --start 2026-04-10T14:00 --end 2026-04-10T15:00 --customer-name "Max Muster" --customer-email max@co.com` |
| `booking-cancel \<business_id\> \<booking_id\>` | Cancel appointment | `m365 booking-cancel BIZ_ID APT_ID --reason "Umgeplant"` |

### SharePoint Sites

| Command | Description | Example |
|---------|-------------|---------|
| `sites-list [N]` | List accessible sites (default 20) | `m365 sites-list` |
| `sites-search \<query\>` | Search sites by keyword | `m365 sites-search "Intranet"` |

## Backward-Compatible Aliases

The following legacy command names are still supported:

| Legacy Command | New Equivalent |
|----------------|---------------|
| `send-mail <to> <subject> <body>` | `mail-send --to ... --subject ... --body ...` |
| `mail-headers <id>` | `mail-read <id> --read-header` |
| `mail-reply-all <id> <body>` | `mail-reply <id> --body ... --reply-all` |
| `mail-delete <id>` | `mail-handle <id> --delete` |
| `mail-move <id> <folder>` | `mail-handle <id> --move <folder>` |
| `contacts-list [N]` | `contact-list [N]` |
| `contacts-get <name>` | `contact-read <id>` |
| `contacts-create <first> <last>` | `contact-create <first> <last>` |
| `contacts-photo-set <id> <file>` | `contact-photo <id> --upload <file>` |
| `contacts-photo-delete <id>` | `contact-photo <id> --delete` |
| `contacts-photo-get <id> <path>` | `contact-photo <id> --download <path>` |
| `upload <local> <remote>` | `onedrive-upload <local> <remote>` |
| `download <remote> <local>` | `onedrive-download <remote> <local>` |
| `todo-list-lists` | `tasklist-list` |
| `todo-create-list <name>` | `tasklist-create <name>` |
| `todo-list-tasks <id>` | `task-list --list-id <id>` |
| `todo-all-tasks` | `task-list` |
| `todo-task-today` | `task-list --status notStarted` |
| `todo-create-task <list-id> <title>` | `task-create --title ... --list-id ...` |
| `todo-update-task <list-id> <task-id>` | `task-update <task-id> --list-id ...` |
| `todo-complete-task <list-id> <task-id>` | `task-update <task-id> --complete --list-id ...` |
| `todo-add-step <list-id> <task-id> <title>` | *(still supported directly)* |
| `todo-complete-step <list-id> <task-id> <step-id>` | *(still supported directly)* |
| `todo-move-task <src> <task-id> <dst>` | *(still supported directly)* |

## Notes
- Authentication is a hybrid model: delegated (device-code) for most features;
  application (client-credentials) for `mail-send` only
- `mail-send` requires `CLIENT_SECRET` and `MAIL_SENDER_UPN` in `.env` and the
  `Mail.Send` application permission granted in Entra ID
- Run `m365 auth-login` for initial setup or after token expiry
- Tokens are cached and auto-refreshed; re-auth needed only after ~90 days of inactivity
- All read/manage API calls use `/me/…` endpoints;
  mail sending uses `/users/{UPN}/sendMail`
- Chat and Meetings require a Microsoft 365 license that includes Teams
- Bookings requires a Microsoft Bookings license in the tenant
- `--count` flag is available on mail-list, contact-list, and task-list
- `--id-only` flag on contact-list prints only IDs (one per line), useful for scripting
- `--short` flag on contact-list prints ID, First Name, Last Name and Work Phone for compact view
- `onedrive-share --anyone` creates an anonymous link; omitting `--anyone` creates an org link
- Excel update uses Graph API directly; file must not be open in Office
- Word and PowerPoint use download → edit locally → re-upload workflow
