# OpenClaw M365 Graph Skill

🦞 **Headless Microsoft 365 integration for OpenClaw agents**  
Control Mail, Calendar, OneNote, OneDrive, Word, Excel and PowerPoint via Microsoft Graph API – fully unattended, daemon-style, Linux-friendly.

This skill gives OpenClaw agents production-ready access to Microsoft 365 services using the **python-o365** library + MSAL with automatic token refresh.  
It supports both **application permissions** (for Mail, Calendar, Files) and **delegated flow** (required for OneNote since early 2025).

### Features
- Send & receive emails (including to external recipients)
- Create, read, update and delete calendar events
- Create and edit OneNote notebooks, sections and pages (HTML content)
- Full OneDrive file operations (list, upload, download, share)
- **Excel** – direct cell/range/table editing via Graph API
- **Word** – download → edit with `python-docx` → re-upload
- **PowerPoint** – download → edit with `python-pptx` → re-upload
- Automatic silent token refresh (persists across reboots / cron jobs)
- Simple CLI: `m365 send-mail ...`, `m365 excel-update ...`, etc.
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

### Microsoft Configuration (Entra ID / Azure AD App Registration) – One-time setup

1. Go to: https://entra.microsoft.com → **App registrations** → **New registration**

   - **Name:** `OpenClaw-M365-Agent` (or similar)  
   - **Supported account types:** Accounts in this organizational directory only (single tenant)  
   - **Redirect URI:** leave blank (daemon / public client)

2. **API permissions** → **Add a permission** → **Microsoft Graph**

   **Application permissions** (for unattended / daemon use):  
   - `Mail.ReadWrite.All`  
   - `Mail.Send`  
   - `Calendars.ReadWrite.All`  
   - `Files.ReadWrite.All`
   - `Contacts.ReadWrite`

   **Delegated permissions** (required for OneNote):  
   - `Notes.ReadWrite.All`  
   - `offline_access`

   → **Grant admin consent for [your organization]**

3. **Certificates & secrets** → **New client secret**

   - **Description:** OpenClaw daemon secret  
   - **Expires:** 24 months (recommended)  
   - **Copy the Value** immediately (you won’t see it again)

4. Copy these three values:

   - **Application (client) ID** → `CLIENT_ID`  
   - **Directory (tenant) ID** → `TENANT_ID`  
   - **Client Secret Value** → `CLIENT_SECRET`

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

# List recent inbox messages
m365 mail-list
m365 mail-list 50

# List upcoming calendar events (next 7 days by default)
m365 calendar-list

# Create a calendar event (ISO 8601 dates)
m365 calendar-create "Team Sync" "2026-04-15T10:00:00" "2026-04-15T11:00:00" "Monthly sync"

# List contacts
m365 contacts-list

# Add a contact
m365 contacts-create "Jane" "Doe" "jane.doe@example.com" "+1-555-0100"

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
# Format: m365 excel-update <path> <sheet> <range> <value1> <value2> ...
m365 excel-update "Reports/Q1-sales.xlsx" Sheet1 A1:B2 "Product" "Revenue" "Total" "124500"

# Replace text placeholders in a Word document
# Format: m365 word-update <path> key1=value1 key2=value2 ...
m365 word-update "Proposals/Offer-2026.docx" "{{Client}}=ACME Corp" "{{Price}}=€ 24,900" "{{Date}}=April 15, 2026"

# Update text in a PowerPoint slide
# Format: m365 ppt-update <path> <slide-number> key1=value1 ...
m365 ppt-update "Presentations/Strategy-2026.pptx" 0 "{{Title}}=2026 Growth Strategy" "{{Subtitle}}=Q2–Q4 Outlook"
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
- Restrict app permissions via Application Access Policy if possible
