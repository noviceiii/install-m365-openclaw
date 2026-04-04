# SKILL: m365-graph

OpenClaw skill for unattended Microsoft 365 access via Graph API.

**Version:** 0.4.0 (RBAC for Applications – Microsoft Recommended)

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
`Files.ReadWrite.All`, `Contacts.ReadWrite`, `Tasks.ReadWrite.All`,
`User.ReadWrite.All`

**Delegated permissions** (required for OneNote):
`Notes.ReadWrite.All`, `offline_access`

## Commands

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
| `mail-send --to ADDR --subject TEXT --body TEXT` | Send an e-mail | `m365 mail-send --to boss@co.com --subject "Update" --body "Done"` |
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
| `calendar-list [N]` | List upcoming events (N = days, default 7) | `m365 calendar-list 14` |
| `calendar-create --subject TEXT --start ISO --end ISO` | Create event | `m365 calendar-create --subject Meeting --start 2026-04-10T10:00 --end 2026-04-10T11:00` |
| `calendar-create ... --body TEXT` | Add description | `m365 calendar-create --subject S --start T1 --end T2 --body "Q1 review"` |
| `calendar-create ... --location TEXT` | Set location | `m365 calendar-create --subject S --start T1 --end T2 --location "Room 101"` |
| `calendar-create ... --required EMAIL` | Add required attendee | `m365 calendar-create --subject S --start T1 --end T2 --required alice@co.com` |
| `calendar-create ... --optional EMAIL` | Add optional attendee | `m365 calendar-create --subject S --start T1 --end T2 --optional bob@co.com` |
| `calendar-create ... --private` | Mark as private | `m365 calendar-create --subject S --start T1 --end T2 --private` |
| `calendar-create ... --reminder N` | Set reminder (minutes) | `m365 calendar-create --subject S --start T1 --end T2 --reminder 30` |
| `calendar-create ... --attach FILE` | Attach file | `m365 calendar-create --subject S --start T1 --end T2 --attach agenda.pdf` |
| `calendar-read \<id\>` | Show event details | `m365 calendar-read AAES...` |
| `calendar-read \<id\> --participants` | List attendees and status | `m365 calendar-read AAES... --participants` |
| `calendar-read \<id\> --cancel` | Cancel event | `m365 calendar-read AAES... --cancel` |
| `calendar-read \<id\> --delete` | Delete event | `m365 calendar-read AAES... --delete` |
| `calendar-handle \<id\> --confirm-accept` | Accept invitation | `m365 calendar-handle AAES... --confirm-accept` |
| `calendar-handle \<id\> --confirm-tentative` | Tentatively accept | `m365 calendar-handle AAES... --confirm-tentative` |
| `calendar-handle \<id\> --confirm-deny` | Decline invitation | `m365 calendar-handle AAES... --confirm-deny` |

### Contacts

| Command | Description | Example |
|---------|-------------|---------|
| `contact-list [N]` | List N contacts (default 100) | `m365 contact-list` |
| `contact-list --sort by-last\|by-first` | Sort contacts | `m365 contact-list --sort by-last` |
| `contact-list --count` | Count contacts | `m365 contact-list --count` |
| `contact-read \<id\>` | Show all fields of a contact | `m365 contact-read AAMk...` |
| `contact-edit \<id\> [--field VALUE ...]` | Update contact fields | `m365 contact-edit AAMk... --email-business neu@domain.ch` |
| `contact-create \<first\> \<last\> [options]` | Create new contact | `m365 contact-create Jane Doe --email-business j@work.com` |
| `contact-photo \<id\> --upload FILE` | Upload contact photo | `m365 contact-photo AAMk... --upload foto.jpg` |
| `contact-photo \<id\> --delete` | Delete contact photo | `m365 contact-photo AAMk... --delete` |
| `contact-photo \<id\> --download PATH` | Download contact photo | `m365 contact-photo AAMk... --download /tmp/foto.jpg` |
| `contactlist-list` | List contact folders | `m365 contactlist-list` |
| `contactlist-create \<name\>` | Create contact folder | `m365 contactlist-create Kunden2026` |
| `contactlist-delete \<id\>` | Delete contact folder | `m365 contactlist-delete AAMk...` |

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
| `todo-rename-list <id> <name>` | *(use tasklist operations)* |
| `todo-delete-list <id>` | *(use tasklist operations)* |
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
- Mail access requires Exchange Online RBAC for Applications (recommended) in addition to Mail.ReadWrite.All and Mail.Send
- Legacy Application Access Policy (New-ApplicationAccessPolicy) is deprecated – use RBAC for Applications instead
- Tasks.ReadWrite.All permission required for ToDo/task features (add to app registration)
- OneNote requires delegated Notes.ReadWrite.All permission
- User.ReadWrite.All is required for user-read and user-update commands
- Excel update uses Graph API directly; file must not be open in Office
- Word and PowerPoint use download → edit locally → re-upload workflow
- `--count` flag is available on mail-list, contact-list, and task-list
- `onedrive-share --anyone` creates an anonymous link; omitting `--anyone` creates an organization link
