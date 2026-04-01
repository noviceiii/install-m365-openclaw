# OpenClaw M365 Graph Skill

🦞 **Headless Microsoft 365 integration for OpenClaw agents**  
Control Mail, Calendar, OneNote, OneDrive, Word, Excel and PowerPoint via Microsoft Graph API – fully unattended, daemon-style, Linux-friendly.

This skill gives OpenClaw agents production-ready access to Microsoft 365 services using the **python-o365** library + MSAL with automatic token refresh.  
It supports both **application permissions** (for Mail, Calendar, Files) and **delegated flow** (required for OneNote since early 2025).

### Features
- Send & receive emails with CC, BCC, attachments, importance, sensitivity, and receipt requests
- Read, reply, reply-all, forward, delete and move messages
- Create, read, update and delete calendar events with attendees, privacy, and reminders
- Create and edit OneNote notebooks, sections and pages (HTML content)
- Full OneDrive file operations (list, upload, download, share)
- Contacts management: create, list, get, photo upload/delete with full field support
- **Excel** – direct cell/range/table editing via Graph API
- **Word** – download → edit with `python-docx` → re-upload
- **PowerPoint** – download → edit with `python-pptx` → re-upload
- **Microsoft To Do** – full task management: lists, tasks, steps, assign, move
- Automatic silent token refresh (persists across reboots / cron jobs)
- Simple CLI: `m365 send-mail ...`, `m365 todo-task-create ...`, etc.
- OpenClaw-native: installs under `~/.openclaw/skills/m365-graph` with `SKILL.md` discovery
- Secure: no hardcoded secrets, encrypted MSAL token cache, `.env` file

### Requirements
- **Microsoft 365 license**  
  Minimum: **Microsoft 365 Business Basic** (or Business Standard / E3 / E5)  
  → Personal / Family plans **do NOT** support application permissions / daemon apps

- Ubuntu 24.04 LTS (or newer) – headless server recommended

- Python 3.10+ (installed automatically)

### Installation (one command)

```bash
curl -sSL https://raw.githubusercontent.com/noviceiii/install-m365-openclaw/main/install-m365-openclaw.sh | bash
```
Or clone the repo and run the installer manually:
```bash
git clone https://github.com/noviceiii/openclaw-m365-graph-skill.git
cd openclaw-m365-graph-skill
chmod +x install-m365-openclaw.sh
./install-m365-openclaw.sh
```
The installer will:

- Create isolated virtual environment
- Install all dependencies (`O365`, `msal`, `python-docx`, `python-pptx`, `openpyxl`, …)
- Create CLI command `m365`
- Set up directory structure under `~/.openclaw/skills/m365-graph`
- Generate `SKILL.md` for automatic OpenClaw agent discovery

### Microsoft Configuration – One-time setup

> For a detailed, step-by-step walkthrough see:
> - [Microsoft-ENTRA-ID-installation.md](Microsoft-ENTRA-ID-installation.md) – App registration & permissions
> - [Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md) – Restrict mailbox access via Exchange Online Application Access Policy (or modern RBAC for Applications)

### Entra ID / Azure AD App Registration – Overview

Register a **daemon application** in Microsoft Entra ID and grant the following Microsoft Graph permissions:

**Application permissions** (unattended/daemon access):  
`Mail.ReadWrite.All`, `Mail.Send`, `Calendars.ReadWrite.All`, `Files.ReadWrite.All`, `Contacts.ReadWrite`, `Tasks.ReadWrite`

**Delegated permissions** (required for OneNote):  
`Notes.ReadWrite.All`, `offline_access`

After registration, create a **client secret** and note down your `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET`.

→ For the full step-by-step guide see [Microsoft-ENTRA-ID-installation.md](Microsoft-ENTRA-ID-installation.md)

### Exchange Online Access Policy – Overview

Exchange Online requires an explicit access policy before the app can read or send mail via the Graph API.
Use the included `setup-exchange-policy.ps1` script to configure this, or follow the manual steps.

→ For the full setup guide see [Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md)

### After Installation – Enter Credentials
The installer will prompt you to enter:
```text
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=~abcdefghijklmnopqrstuvwxyz1234567890abcdef
```

Alternatively edit the file manually:
```bash
nano ~/.openclaw/skills/m365-graph/.env
```

### First Run & Token Authentication
Run once to authenticate (only needed the very first time):
```bash
m365
```
→ Browser window opens (or copy-paste URL if headless)
→ Sign in with an admin / licensed user account
→ Consent (only once)
→ Token is saved and will auto-refresh forever

### CLI Examples

After installation and successful authentication, your OpenClaw agents (or you manually) can use the `m365` command like this:

```bash
# Send an email (works to external recipients too)
m365 send-mail colleague@external-company.com "Project Update" "Please find the attached Q3 report. Deadline is Friday."

# Send with CC, high importance, and attachment
m365 send-mail boss@company.com "Q3 Report" "See attached" --cc cfo@company.com --importance High --attach ./report.pdf

# List recent inbox messages
m365 mail-list
m365 mail-list 50

# Read messages by subject
m365 mail-read-subject "Invoice"

# Read, reply, reply-all, forward, delete, move a message
m365 mail-read AAMkAGI...
m365 mail-reply AAMkAGI... "Got it, thanks!"
m365 mail-reply-all AAMkAGI... "All noted, proceeding."
m365 mail-forward AAMkAGI... colleague@company.com "FYI"
m365 mail-delete AAMkAGI...
m365 mail-move AAMkAGI... Archive

# List upcoming calendar events (next 7 days by default)
m365 calendar-list

# Create a calendar event (ISO 8601 dates)
m365 calendar-create "Team Sync" "2026-04-15T10:00:00" "2026-04-15T11:00:00" "Monthly sync"

# Create an event with attendees, location, and privacy
m365 calendar-create "Budget Review" "2026-04-20T14:00:00" "2026-04-20T15:00:00" \
  --required alice@company.com --optional bob@company.com \
  --location "Board Room" --private --reminder-minutes 30

# List contacts
m365 contacts-list

# Add a contact (simple)
m365 contacts-create "Jane" "Doe" "jane.doe@example.com" "+1-555-0100"

# Add a contact (extended)
m365 contacts-create "Jane" "Doe" \
  --email-business jane@work.com --email-personal jane@home.com \
  --phone-business +1-555-0100 --phone-mobile +1-555-0101 \
  --birthday 1985-06-15 --notes "Met at conference 2024"

# Get all fields of a contact
m365 contacts-get AAMkAGI...

# Update or delete contact photo
m365 contacts-photo-update AAMkAGI... ./photo.jpg
m365 contacts-photo-delete AAMkAGI...

# Create a OneNote page (notebook, section, title, html)
m365 onenote-create "WorkNotes" "April 2026" "Team Sync" "<h1>Team Sync April</h1><p>Agenda: Budget review, new hires</p>"

# List files in OneDrive root or a folder
m365 onedrive-list "/"
m365 onedrive-list "/Documents/Reports"

# Upload a local file to OneDrive
m365 upload ./budget-2026.xlsx "/Finance/Annual/Budget 2026.xlsx"

# Download a file from OneDrive
m365 download "/Finance/Annual/Budget 2026.xlsx" ./local-budget.xlsx

# Update cells in an Excel file (direct Graph API – no download needed)
m365 excel-update "Reports/Q1-sales.xlsx" Sheet1 A1:B2 "Product" "Revenue" "Total" "124500"

# Replace text placeholders in a Word document
m365 word-update "Proposals/Offer-2026.docx" "{{Client}}=ACME Corp" "{{Price}}=€ 24,900" "{{Date}}=April 15, 2026"

# Update text in a PowerPoint slide
m365 ppt-update "Presentations/Strategy-2026.pptx" 0 "{{Title}}=2026 Growth Strategy" "{{Subtitle}}=Q2–Q4 Outlook"

# To Do: list task lists and tasks
m365 todo-lists
m365 todo-tasks AAMk...
m365 todo-tasks-all --due-before 2026-05-01

# To Do: create, update, complete, delete tasks
m365 todo-task-create AAMk... "Prepare report" --due 2026-04-15 --note "Include Q1 data"
m365 todo-task-update AAMk... AAMk2... --title "Prepare final report"
m365 todo-task-complete AAMk... AAMk2...
m365 todo-task-delete AAMk... AAMk2...

# To Do: steps, assign, move
m365 todo-step-add AAMk... AAMk2... "Collect data"
m365 todo-step-complete AAMk... AAMk2... AAMk3...
m365 todo-task-assign AAMk... AAMk2... alice@company.com
m365 todo-task-move AAMk... AAMk2... AAMk4...
```

## Troubleshooting
- Token expired / authentication fails → delete ~/.openclaw/credentials/m365_token_cache.bin and run m365 again
- Permission denied → check admin consent was granted
- Email not arriving externally → configure SPF/DKIM/DMARC in your Microsoft 365 tenant
- Excel/PowerPoint editing fails → make sure file is not open in desktop Office app

## Security Notes
- Never commit .env or the token cache to git
- Use a dedicated service account with minimal licenses
- Rotate client secret every 12–24 months
- Restrict app permissions via Application Access Policy (or modern RBAC for Applications) – see [Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md)
