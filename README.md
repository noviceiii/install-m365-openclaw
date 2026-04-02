# install-m365-openclaw

OpenClaw Skill für unattended Microsoft 365 Zugriff via Graph API.

**Version:** 0.3.0 (RBAC for Applications – Microsoft Recommended)  
**Ziel:** Cleotine Claw als persönliche Assistentin im Multiuser-OpenClaw-System auf Ubuntu Server (headless).

## Schnellstart

```bash
curl -L https://raw.githubusercontent.com/noviceiii/install-m365-openclaw/main/install-m365-openclaw.sh | bash
```

## Voraussetzungen & Setup

1. Entra ID App Registration → [Microsoft-ENTRA-ID-installation.md](Microsoft-ENTRA-ID-installation.md)
2. Exchange Online RBAC (Pflicht für Mail) → [Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md)
3. Credentials in `~/.openclaw/skills/m365-graph/.env` eintragen
4. Erste Auth: `m365 calendar-list`

Siehe [SKILL.md](SKILL.md) für alle verfügbaren Befehle.
