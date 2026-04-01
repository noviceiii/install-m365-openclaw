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
