#!/usr/bin/env bash
# =============================================================================
# OpenClaw M365 Graph Skill - Installer 
# =============================================================================
# Improvements:
# - Asks for OpenClaw base directory
# - Asks for M365 user email (guidance only)
# - Creates proper package structure: src/m365_openclaw/__init__.py
# - Installs as editable package with correct setup.py
# - Verifies package import after installation
# - PATH added permanently (only once)
# - Uses venv Python and pip explicitly
# =============================================================================

set -euo pipefail

# Default values
DEFAULT_OPENCLAW_DIR="${HOME}/.openclaw"
DEFAULT_M365_USER="cleotine.claw@t9t.ch"

echo "=== OpenClaw M365 Graph Skill Installer - Reliable Version ==="

# Ask for OpenClaw base directory
read -p "OpenClaw base directory [default: ${DEFAULT_OPENCLAW_DIR}]: " OPENCLAW_DIR
OPENCLAW_DIR="${OPENCLAW_DIR:-${DEFAULT_OPENCLAW_DIR}}"

# Ask for M365 user email (guidance only)
read -p "M365 account email for authentication [default: ${DEFAULT_M365_USER}]: " M365_USER
M365_USER="${M365_USER:-${DEFAULT_M365_USER}}"

SKILL_NAME="m365-graph"
INSTALL_DIR="${OPENCLAW_DIR}/skills/${SKILL_NAME}"
VENV_DIR="${INSTALL_DIR}/venv"
CLI_NAME="m365"
CONFIG_FILE="${INSTALL_DIR}/.env"
TOKEN_CACHE="${OPENCLAW_DIR}/credentials/m365_token_cache.bin"
SRC_DIR="${INSTALL_DIR}/src"

echo ""
echo "Installation summary:"
echo "  OpenClaw base:     ${OPENCLAW_DIR}"
echo "  Install path:      ${INSTALL_DIR}"
echo "  Venv path:         ${VENV_DIR}"
echo "  Suggested account: ${M365_USER}"
echo ""

# 1. System dependencies (sudo required)
sudo apt update -qq
sudo apt install -y python3 python3-venv python3-pip curl git

# 2. Create directories
mkdir -p "${INSTALL_DIR}"/{bin,src/m365_openclaw}
mkdir -p "${OPENCLAW_DIR}/credentials"
mkdir -p "${HOME}/.local/bin"

# 3. Create virtual environment and install packages
python3 -m venv "${VENV_DIR}"
"${VENV_DIR}/bin/pip" install --upgrade pip
"${VENV_DIR}/bin/pip" install \
    O365 msal msal_extensions python-dotenv \
    python-docx python-pptx openpyxl

# 4. Create proper package structure and setup.py
touch "${SRC_DIR}/m365_openclaw/__init__.py"

cat > "${INSTALL_DIR}/setup.py" << 'EOF'
from setuptools import setup, find_packages

setup(
    name='m365_openclaw',
    version='0.1.0',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
)
EOF

# Install editable package
cd "${INSTALL_DIR}"
echo "Installing m365_openclaw package in editable mode..."
"${VENV_DIR}/bin/pip" install -e .

# Verify import works
echo "Verifying package import..."
if "${VENV_DIR}/bin/python" -c "import m365_openclaw; print('SUCCESS: m365_openclaw imported')" 2>/dev/null; then
    echo "Package verification successful."
else
    echo "ERROR: Package import failed. Check setup.py and src/m365_openclaw/__init__.py"
    exit 1
fi

# 5. Create CLI wrapper
cat > "${INSTALL_DIR}/bin/m365" << EOF
#!/usr/bin/env bash
export PYTHONPATH="${SRC_DIR}:\${PYTHONPATH}"
exec "${VENV_DIR}/bin/python" -m m365_openclaw "\$@"
EOF

chmod +x "${INSTALL_DIR}/bin/m365"

ln -sf "${INSTALL_DIR}/bin/m365" "${HOME}/.local/bin/${CLI_NAME}"

# Add PATH permanently (only once)
if ! grep -q "${HOME}/.local/bin" "${HOME}/.bashrc"; then
    echo 'export PATH="$HOME/.local/bin:$PATH"' >> "${HOME}/.bashrc"
    echo "Added PATH entry to ~/.bashrc"
fi
export PATH="${HOME}/.local/bin:$PATH"

# 6. Create config file
cat > "${CONFIG_FILE}" << EOF
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-app-client-id-here
CLIENT_SECRET=your-app-client-secret-here
TOKEN_CACHE_PATH=${TOKEN_CACHE}
EOF

# 7. Entra ID instructions
echo ""
echo "=== Entra ID App Registration (one-time) ==="
echo "1. Go to https://entra.microsoft.com → App registrations → New registration"
echo "   Name: OpenClaw-M365-Agent"
echo "   Supported account types: Accounts in this organizational directory only"
echo "   Redirect URI: leave blank"
echo "2. API permissions → Microsoft Graph"
echo "   Application: Mail.ReadWrite.All, Mail.Send, Calendars.ReadWrite.All, Files.ReadWrite.All"
echo "   Delegated:   Notes.ReadWrite.All, offline_access"
echo "   → Grant admin consent"
echo "3. Certificates & secrets → New client secret → copy Value"
echo ""
read -p "Press ENTER when you have TENANT_ID, CLIENT_ID, CLIENT_SECRET..."

read -p "TENANT_ID: " TENANT_ID
read -p "CLIENT_ID: " CLIENT_ID
read -s -p "CLIENT_SECRET: " CLIENT_SECRET
echo ""

sed -i "s|your-tenant-id-here|${TENANT_ID}|" "${CONFIG_FILE}"
sed -i "s|your-app-client-id-here|${CLIENT_ID}|" "${CONFIG_FILE}"
sed -i "s|your-app-client-secret-here|${CLIENT_SECRET}|" "${CONFIG_FILE}"

# 8. Create modules inside package
cat > "${SRC_DIR}/m365_openclaw/core.py" << 'EOF'
# core.py - Microsoft 365 Client
import os
from pathlib import Path
from dotenv import load_dotenv
from O365 import Account
from O365.utils import FileSystemTokenBackend
from datetime import datetime, timedelta, timezone

load_dotenv()

class M365Client:
    def __init__(self):
        self.tenant_id = os.getenv("TENANT_ID")
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.credentials = (self.client_id, self.client_secret)
        
        self.token_backend = FileSystemTokenBackend(
            token_path=Path(os.getenv("TOKEN_CACHE_PATH"))
        )
        
        self.account = Account(
            credentials=self.credentials,
            auth_flow_type='credentials',
            tenant_id=self.tenant_id,
            token_backend=self.token_backend
        )
        
        if not self.account.is_authenticated:
            print("Initial authentication required...")
            self.account.authenticate(
                scopes=['https://graph.microsoft.com/.default']
            )
        
        self.mailbox = self.account.mailbox()
        self.calendar = self.account.schedule()
        self.drive = self.account.storage()

    def get_calendar_events(self, days=7):
        try:
            events = self.calendar.get_events(include_recurring=False, limit=100)
            now_utc = datetime.now(timezone.utc)
            cutoff = now_utc + timedelta(days=days)
            upcoming = []
            for event in events:
                start_time = event.start
                if start_time.tzinfo is None:
                    start_time = start_time.replace(tzinfo=timezone.utc)
                if start_time >= now_utc and start_time <= cutoff:
                    upcoming.append({
                        'subject': event.subject or '(no subject)',
                        'start': start_time.isoformat(),
                        'end': event.end.isoformat() if event.end else None
                    })
            return upcoming
        except Exception as e:
            return f"Error: {str(e)}"
EOF

cat > "${SRC_DIR}/m365_openclaw/__main__.py" << 'EOF'
# __main__.py - CLI entry point
import sys
from m365_openclaw.core import M365Client

client = M365Client()

if len(sys.argv) < 2:
    print("OpenClaw M365 CLI ready!")
    print("Usage: m365 <command> [args]")
    sys.exit(0)

cmd = sys.argv[1]

if cmd == "calendar-list":
    events = client.get_calendar_events()
    if isinstance(events, str):
        print(events)
    elif not events:
        print("No events in the next 7 days.")
    else:
        print("Upcoming events:")
        for e in events:
            print(f"- {e['start']} | {e['subject']}")
else:
    print(f"Unknown command: {cmd}")
EOF

# 9. Start initial authentication
echo ""
echo "[9/10] Starting initial authentication..."
echo "Please sign in with: ${M365_USER}"
echo "If browser does not open (headless system), copy the URL from terminal."
echo ""

export PYTHONPATH="${SRC_DIR}:${PYTHONPATH}"
"${VENV_DIR}/bin/python" -m m365_openclaw calendar-list

echo ""
echo "=== Installation complete! ==="
echo "Test with:"
echo "  m365 calendar-list"
echo ""
echo "If authentication fails or no calendar appears:"
echo "  rm -rf ${OPENCLAW_DIR}/credentials/m365_token_cache.bin"
echo "  cd ${INSTALL_DIR}"
echo "  ${VENV_DIR}/bin/python -m m365_openclaw calendar-list"
echo ""
echo "Done!"
