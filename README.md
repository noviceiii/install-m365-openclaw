# install-m365-openclaw

OpenClaw Skill for Microsoft 365 access via Graph API with **hybrid authentication**:
delegated Device-Code for most features, and a confidential client (`CLIENT_SECRET`)
for application-permission `Mail.Send`.

**It only works with a business m365 licence.**

**Version:** 0.6.1 (Hybrid: Delegated Device-Code + Application Client-Credentials)  
**Goal:** Headless M365 access for OpenClaw agents – one-time sign-in via Device-Code, then permanently token-based. Mail sending uses application permissions.

This README is the documentation master. Command details live in [SKILL.md](SKILL.md).

## Quick Start

```bash
curl -L https://raw.githubusercontent.com/noviceiii/install-m365-openclaw/main/install-m365-openclaw.sh | bash
```

The installer copies `setup.py` and `src/` from the directory next to the script
when those files are present (local clone). If they are missing (typical `curl | bash`
invocation), it clones this repository into a well-known temp directory and continues.

## Features

This skill gives OpenClaw agents full access to your Microsoft 365 account. Key capabilities include:

- **Mail** – Send and read e-mails, reply, forward, manage folders
- **Contacts** – Create, edit, and delete contacts and contact folders
- **Tasks** – Manage Microsoft To Do tasks and task lists
- **Calendar** – Create, read, and manage calendar events; handle invitations (accept, tentative, decline)
- **OneDrive** – Upload, download, move, and share files
- **Teams Chat & Online Meetings** – Create chats, send messages, schedule online meetings
- **OneNote** – Create and list notebook pages
- **Office Documents** – Update Excel, Word, and PowerPoint files via Graph API
- **Bookings** – Manage Microsoft Bookings appointments
- **SharePoint** – List and search SharePoint sites

For a full list of all commands and options, see [SKILL.md](SKILL.md).

## Prerequisites & Setup

1. Entra ID App Registration (hybrid: delegated + application) → [Microsoft-ENTRA-ID-installation.md](Microsoft-ENTRA-ID-installation.md)
   - Enable **Public client / device-code** flow
   - Grant delegated Graph permissions (admin consent)
   - Grant the **Mail.Send application** permission (admin consent)
   - Create a **client secret**
2. Enter credentials in `~/.openclaw/skills/m365-graph/.env`:
   `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `MAIL_SENDER_UPN`
3. Configure Exchange Online RBAC for Applications **before using `mail-send`**:
   run [setup-exchange-policy.ps1](setup-exchange-policy.ps1)
   (see [Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md))
4. First sign-in via Device-Code: `m365 auth-login`
5. After that, run headlessly: `m365 calendar-list`

See [SKILL.md](SKILL.md) for all available commands.

## Authentication

The skill uses a **hybrid** model:

- **Delegated (device-code)** – for all features except mail sending (`GET /me/…`).
  The user signs in **once** interactively at **https://microsoft.com/devicelogin**.
  The code is displayed in the terminal. After sign-in the app stores refresh tokens
  and works **permanently headless**. No browser is required on the server.
- **Application (client-credentials)** – used exclusively for `mail-send`
  (`POST /users/{MAIL_SENDER_UPN}/sendMail`). This requires:
  - `CLIENT_SECRET` and `MAIL_SENDER_UPN` in `.env`
  - the **Mail.Send** *application* permission, admin-consented in Entra ID
  - Exchange RBAC for Applications via `setup-exchange-policy.ps1` **before** first `mail-send`

The same Entra ID app registration is both a public client (device-code) and a
confidential client (client secret).

## OpenClaw Integration

### Skill Registration

The installer automatically registers the skill in `~/.openclaw/openclaw.json`:

```json
{
  "skills": {
    "entries": {
      "m365-graph": {
        "enabled": true
      }
    }
  }
}
```

Existing configuration values are **preserved** – only missing entries are added.

### Multi-Agent Operation

After installation the skill is available as a **shared resource** for all OpenClaw agents
on this machine (via `~/.openclaw/skills/m365-graph`).

To restrict the skill to specific agents, an allowlist can be configured in `~/.openclaw/openclaw.json`:

```json
{
  "agents": {
    "defaults": {
      "skills": ["m365-graph"]
    },
    "list": [
      { "id": "cleotine" },
      { "id": "restricted-agent", "skills": [] }
    ]
  }
}
```

### Skill Manager

`SKILL.md` contains YAML frontmatter according to the [AgentSkills](https://agentskills.io) specification
and is therefore compatible with the OpenClaw Skill Manager:

```bash
# Check skill status (based on openclaw.json configuration)
openclaw skills list
```

### Re-installation / Update

The installer script can be run multiple times:
- **Existing `.env` values** (credentials) are **not overwritten** – only missing keys are added
- **Existing `openclaw.json` entries** are preserved
- Python packages and the CLI wrapper are updated

## License

Copyright © noviceiii

**Private use permitted** – You are free to use and modify this software for personal,
non-commercial purposes.

**Commercial use prohibited** – The use of this software, in whole or in part, for any
commercial purpose (including internal business operations of a for-profit organization)
is **not permitted** without prior written permission from the author.

This software is provided "as is", without warranty of any kind.
