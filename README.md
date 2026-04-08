# install-m365-openclaw

OpenClaw Skill for Microsoft 365 access via Graph API with delegated permissions (Device-Code Flow).

## It only works with a business m365 licence.

**Version:** 0.6.0 (Delegated Permissions – Device-Code Flow - reworked and enhanced command structure)  
**Goal:** Headless M365 access for OpenClaw agents – one-time sign-in via Device-Code, then permanently token-based.

## Quick Start

```bash
curl -L https://raw.githubusercontent.com/noviceiii/install-m365-openclaw/main/install-m365-openclaw.sh | bash
```

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

1. Entra ID App Registration (Delegated Permissions) → [Microsoft-ENTRA-ID-installation.md](Microsoft-ENTRA-ID-installation.md)
2. Enter credentials in `~/.openclaw/skills/m365-graph/.env` (TENANT_ID, CLIENT_ID)
3. First sign-in via Device-Code: `m365 auth-login`
4. After that, run headlessly: `m365 calendar-list`

See [SKILL.md](SKILL.md) for all available commands.

## Authentication

The skill uses **delegated permissions** with **Device-Code Flow**:

- The user signs in **once** interactively at **https://microsoft.com/devicelogin**
- The code is displayed in the terminal
- After sign-in the app stores Refresh Tokens and works **permanently headless**
- No browser required on the server
- No client secret required

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
