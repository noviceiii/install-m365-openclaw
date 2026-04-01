# OpenClaw M365 Graph Skill

🦞 **Headless Microsoft 365 integration for OpenClaw agents**  
Control Mail, Calendar, Contacts, ToDo, OneNote, OneDrive, Word, Excel and PowerPoint via Microsoft Graph API – fully unattended, daemon-style, Linux-friendly.

This skill gives OpenClaw agents production-ready access to Microsoft 365 services using the **python-o365** library + MSAL with automatic token refresh.  
It supports both **application permissions** (for Mail, Calendar, Contacts, Files, Tasks) and **delegated flow** (required for OneNote since early 2025).

### Features
- **Mail**: Send (To/CC/BCC, priority, sensitivity, attachments, receipts), list, search, read headers, reply, reply-all, forward, delete, move
- **Calendar**: Create events with required/optional attendees, location, private flag, reminder, file attachment; list upcoming events
- **Contacts**: Full CRUD with typed emails/phones, home/work addresses, birthday, anniversary, website, spouse, notes, and profile photo management
- **Microsoft ToDo**: Full task list and task management – create/rename/delete lists; create/update/complete/move tasks; checklist steps
- **OneNote**: Create notebooks, sections and pages (HTML content)
- **OneDrive**: List, upload, download, share files
- **Excel**: Direct cell/range/table editing via Graph API (no download)
- **Word**: Download → edit with `python-docx` → re-upload
- **PowerPoint**: Download → edit with `python-pptx` → re-upload
- Automatic silent token refresh (persists across reboots / cron jobs)
- Simple CLI: `m365 send-mail ...`, `m365 todo-create-task ...`, etc.
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
`Mail.ReadWrite.All`, `Mail.Send`, `Calendars.ReadWrite.All`, `Files.ReadWrite.All`, `Contacts.ReadWrite`, `Tasks.ReadWrite.All`

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

#### Mail

```bash
# List recent inbox messages (includes message IDs for follow-up operations)
m365 mail-list
m365 mail-list 50

# Send an email (basic – same as v0.1.0)
m365 send-mail colleague@external.com "Project Update" "Please find the Q3 report attached."

# Send with CC, BCC, priority, file attachment and delivery receipt
m365 send-mail alice@acme.com "Offer" "Please review." \
    --cc manager@acme.com --bcc archive@acme.com \
    --importance High --sensitivity Confidential \
    --attach ./offer.pdf \
    --delivery-receipt --read-receipt

# Search inbox by subject
m365 mail-search "Invoice"

# Read full headers of a message (get the ID from mail-list or mail-search)
m365 mail-headers AAMkAGVm...

# Reply, reply-all, forward
m365 mail-reply   AAMkAGVm... "Thanks, will do!"
m365 mail-reply-all AAMkAGVm... "Noted by everyone."
m365 mail-forward AAMkAGVm... ceo@acme.com "FYI – see below."

# Delete or move a message
m365 mail-delete AAMkAGVm...
m365 mail-move   AAMkAGVm... archive
```

#### Calendar

```bash
# List upcoming calendar events (next 7 days by default)
m365 calendar-list
m365 calendar-list 30

# Create a basic event (same as v0.1.0)
m365 calendar-create "Team Sync" "2026-04-15T10:00:00" "2026-04-15T11:00:00" "Monthly sync"

# Create event with attendees, location, private flag and 15-minute reminder
m365 calendar-create "Board Meeting" "2026-04-20T09:00:00" "2026-04-20T11:00:00" "Q2 Review" \
    --attendees cfo@acme.com,cto@acme.com \
    --optional-attendees pa@acme.com \
    --location "HQ Conference Room A" \
    --private \
    --reminder 15 \
    --attach ./agenda.pdf
```

#### Contacts

```bash
# List all contacts (table: First, Last, Work Mail, Home Mail, Work Phone, Home Phone, Mobile)
m365 contacts-list

# Show all fields for a specific contact
m365 contacts-get "Jane Doe"

# Create a contact (basic – same as v0.1.0)
m365 contacts-create "Jane" "Doe" "jane.doe@example.com" "+1-555-0100"

# Create contact with extended fields
m365 contacts-create "Max" "Mustermann" max@acme.com +49-89-123456 \
    --home-email max@private.de \
    --mobile-phone "+49-172-9876543" \
    --work-city "Munich" --work-country "Germany" \
    --home-city "Augsburg" --home-country "Germany" \
    --birthday "1985-03-22T00:00:00Z" \
    --anniversary "2010-06-15T00:00:00Z" \
    --website "https://max.example.de" \
    --work-website "https://acme.com/team/max" \
    --spouse "Maria Mustermann" \
    --notes "Hobbies: Cycling, Photography. Zodiac: Aries. Children: 2"

# Manage profile photos (use the contact ID from contacts-get)
m365 contacts-photo-set  AAMkAGVm... /tmp/jane.jpg
m365 contacts-photo-get  AAMkAGVm... /tmp/jane_download.jpg
m365 contacts-photo-delete AAMkAGVm...
```

#### Microsoft ToDo

```bash
# List all task lists
m365 todo-list-lists

# Create, rename and delete task lists
m365 todo-create-list "Project Alpha"
m365 todo-rename-list AAMkAGVm... "Project Beta"
m365 todo-delete-list AAMkAGVm...

# List tasks in a specific list (use list ID from todo-list-lists)
m365 todo-list-tasks AAMkAGVm...
m365 todo-list-tasks AAMkAGVm... --due-after 2026-04-01

# List all tasks across every list
m365 todo-all-tasks
m365 todo-all-tasks --due-after 2026-04-01

# Create a task
m365 todo-create-task AAMkAGVm... "Review budget report" \
    --note "Check Q2 numbers against forecast" \
    --due "2026-04-20T09:00:00" \
    --reminder "2026-04-19T08:00:00"

# Update, complete and move tasks
m365 todo-update-task AAMkAGVm... BBMkAGVm... --due "2026-04-25T09:00:00"
m365 todo-complete-task AAMkAGVm... BBMkAGVm...
m365 todo-move-task     FromListId  BBMkAGVm... ToListId

# Checklist steps (subtasks)
m365 todo-add-step      AAMkAGVm... BBMkAGVm... "Draft executive summary"
m365 todo-complete-step AAMkAGVm... BBMkAGVm... CCMkAGVm...
```

#### OneDrive, OneNote, Excel, Word, PowerPoint

```bash
# List files in OneDrive root or a folder
m365 onedrive-list "/"
m365 onedrive-list "/Documents/Reports"

# Upload a local file to OneDrive
m365 upload ./budget-2026.xlsx "/Finance/Annual/Budget 2026.xlsx"

# Download a file from OneDrive
m365 download "/Finance/Annual/Budget 2026.xlsx" ./local-budget.xlsx

# Create a OneNote page (notebook, section, title, html)
m365 onenote-create "WorkNotes" "April 2026" "Team Sync" "<h1>Team Sync April</h1><p>Agenda: Budget review, new hires</p>"

# Update cells in an Excel file (direct Graph API – no download needed)
m365 excel-update "Reports/Q1-sales.xlsx" Sheet1 A1:B2 "Product" "Revenue" "Total" "124500"

# Replace text placeholders in a Word document
m365 word-update "Proposals/Offer-2026.docx" "{{Client}}=ACME Corp" "{{Price}}=€ 24,900" "{{Date}}=April 15, 2026"

# Update text in a PowerPoint slide
m365 ppt-update "Presentations/Strategy-2026.pptx" 0 "{{Title}}=2026 Growth Strategy" "{{Subtitle}}=Q2–Q4 Outlook"
```

## Troubleshooting
- Token expired / authentication fails → delete ~/.openclaw/credentials/m365_token_cache.bin and run m365 again
- Permission denied → check admin consent was granted for all required permissions (incl. Tasks.ReadWrite.All for ToDo)
- Email not arriving externally → configure SPF/DKIM/DMARC in your Microsoft 365 tenant
- Excel/PowerPoint editing fails → make sure file is not open in desktop Office app
- ToDo commands fail with 403 → grant Tasks.ReadWrite.All application permission in Entra ID

## Security Notes
- Never commit .env or the token cache to git
- Use a dedicated service account with minimal licenses
- Rotate client secret every 12–24 months
- Restrict app permissions via Application Access Policy (or modern RBAC for Applications) – see [Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md)
