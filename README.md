# install-m365-openclaw

OpenClaw Skill für Microsoft 365 Zugriff via Graph API mit delegierten Berechtigungen (Device-Code Flow).

**Version:** 0.5.1 (Delegated Permissions – Device-Code Flow)  
**Ziel:** Headless M365-Zugriff für OpenClaw-Agenten – einmalige Anmeldung per Device-Code, danach dauerhaft tokenbasiert.

## Schnellstart

```bash
curl -L https://raw.githubusercontent.com/noviceiii/install-m365-openclaw/main/install-m365-openclaw.sh | bash
```

## Voraussetzungen & Setup

1. Entra ID App Registration (Delegated Permissions) → [Microsoft-ENTRA-ID-installation.md](Microsoft-ENTRA-ID-installation.md)
2. Credentials in `~/.openclaw/skills/m365-graph/.env` eintragen (TENANT_ID, CLIENT_ID)
3. Erste Anmeldung per Device-Code: `m365 auth-login`
4. Danach headless: `m365 calendar-list`

Siehe [SKILL.md](SKILL.md) für alle verfügbaren Befehle.

## Authentifizierung

Die Skill verwendet **delegierte Berechtigungen** mit **Device-Code Flow**:

- Der Benutzer meldet sich **einmalig** interaktiv an auf **https://microsoft.com/devicelogin**
- Der Code wird im Terminal angezeigt
- Nach der Anmeldung speichert die App Refresh-Tokens und arbeitet **dauerhaft headless**
- Kein Browser auf dem Server nötig
- Kein Client-Secret nötig

## OpenClaw Integration

### Skill-Registrierung

Der Installer registriert den Skill automatisch in `~/.openclaw/openclaw.json`:

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

Bestehende Konfigurationswerte werden dabei **erhalten** – nur fehlende Einträge werden ergänzt.

### Multi-Agent Betrieb

Nach der Installation ist der Skill als **geteilte Ressource** für alle OpenClaw-Agenten
auf dieser Maschine verfügbar (via `~/.openclaw/skills/m365-graph`).

Um den Skill auf bestimmte Agenten zu beschränken, kann in `~/.openclaw/openclaw.json`
eine Allowlist konfiguriert werden:

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

### Skill-Manager

Das `SKILL.md` enthält YAML-Frontmatter gemäss der [AgentSkills](https://agentskills.io)-Spezifikation
und ist damit kompatibel mit dem OpenClaw Skill-Manager:

```bash
# Skill-Status prüfen (nach openclaw.json Konfiguration)
openclaw skills list
```

### Erneute Installation / Update

Das Installer-Script kann mehrfach ausgeführt werden:
- **Bestehende `.env` Werte** (Credentials) werden **nicht überschrieben** – nur fehlende Keys werden ergänzt
- **Bestehende `openclaw.json` Einträge** bleiben erhalten
- Python-Pakete und der CLI-Wrapper werden aktualisiert
