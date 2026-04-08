#!/usr/bin/env bash
# =============================================================================
# OpenClaw M365 Graph Skill Installer v0.5.1
# =============================================================================
set -euo pipefail

DEFAULT_OPENCLAW_DIR="${HOME}/.openclaw"

echo "=== OpenClaw M365 Graph Skill Installer v0.5.1 ==="

read -p "OpenClaw base directory [${DEFAULT_OPENCLAW_DIR}]: " OPENCLAW_DIR
OPENCLAW_DIR="${OPENCLAW_DIR:-${DEFAULT_OPENCLAW_DIR}}"

SKILL_NAME="m365-graph"
INSTALL_DIR="${OPENCLAW_DIR}/skills/${SKILL_NAME}"
VENV_DIR="${INSTALL_DIR}/venv"
TOKEN_CACHE="${OPENCLAW_DIR}/credentials/m365_token_cache.bin"
CLI_NAME="m365"

sudo apt update -qq && sudo apt install -y python3 python3-venv python3-pip git

mkdir -p "${INSTALL_DIR}"/{src/m365_openclaw,bin} "${OPENCLAW_DIR}/credentials" "${HOME}/.local/bin"

# venv + dependencies
python3 -m venv "${VENV_DIR}"
"${VENV_DIR}/bin/pip" install --upgrade pip
"${VENV_DIR}/bin/pip" install msal python-dotenv python-docx python-pptx openpyxl requests

# Copy package files from repository structure
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cp "${SCRIPT_DIR}/setup.py" "${INSTALL_DIR}/"
cp -r "${SCRIPT_DIR}/src/m365_openclaw/." "${INSTALL_DIR}/src/m365_openclaw/"

# Copy SKILL.md so OpenClaw's skill manager can read it from the install dir
cp "${SCRIPT_DIR}/SKILL.md" "${INSTALL_DIR}/SKILL.md"

cd "${INSTALL_DIR}"
"${VENV_DIR}/bin/pip" install -e .

# CLI wrapper
cat > "${INSTALL_DIR}/bin/m365" << EOF
#!/usr/bin/env bash
export PYTHONPATH="${INSTALL_DIR}/src:\${PYTHONPATH:-}"
exec "${VENV_DIR}/bin/python" -m m365_openclaw "\$@"
EOF
chmod +x "${INSTALL_DIR}/bin/m365"
ln -sf "${INSTALL_DIR}/bin/m365" "${HOME}/.local/bin/${CLI_NAME}"

if ! grep -q "${HOME}/.local/bin" "${HOME}/.bashrc"; then
    echo 'export PATH="$HOME/.local/bin:$PATH"' >> "${HOME}/.bashrc"
fi

# .env – preserve existing values, only add missing keys
# CLIENT_SECRET and MAIL_SENDER_UPN are required for mail sending via
# application permissions (POST /users/{UPN}/sendMail).
_set_env_if_missing() {
    local key="$1" default="$2" file="$3"
    if ! grep -q "^${key}=" "${file}" 2>/dev/null; then
        echo "${key}=${default}" >> "${file}"
    fi
}
ENV_FILE="${INSTALL_DIR}/.env"
if [ ! -f "${ENV_FILE}" ]; then
    # Create the file with placeholder template values
    cat > "${ENV_FILE}" << EOF
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-app-client-id-here
# Required for application-permission mail sending (client-credentials flow):
CLIENT_SECRET=your-client-secret-here
MAIL_SENDER_UPN=sender@yourdomain.com
TOKEN_CACHE_PATH=${TOKEN_CACHE}
EOF
else
    # File already exists – add only keys that are not yet present
    _set_env_if_missing "TENANT_ID"        "your-tenant-id-here"     "${ENV_FILE}"
    _set_env_if_missing "CLIENT_ID"        "your-app-client-id-here" "${ENV_FILE}"
    _set_env_if_missing "CLIENT_SECRET"    "your-client-secret-here" "${ENV_FILE}"
    _set_env_if_missing "MAIL_SENDER_UPN"  "sender@yourdomain.com"   "${ENV_FILE}"
    _set_env_if_missing "TOKEN_CACHE_PATH" "${TOKEN_CACHE}"           "${ENV_FILE}"
fi

# Register the skill in ~/.openclaw/openclaw.json
# Preserves all existing config; only adds/enables the m365-graph entry.
OPENCLAW_CONFIG="${OPENCLAW_DIR}/openclaw.json"
python3 - "${OPENCLAW_CONFIG}" << 'PYEOF'
import json, os, sys

config_file = sys.argv[1]
config = {}
if os.path.isfile(config_file):
    try:
        with open(config_file) as f:
            config = json.load(f)
    except (json.JSONDecodeError, OSError):
        config = {}

# Ensure skills → entries → m365-graph is present and enabled.
skills  = config.setdefault("skills", {})
entries = skills.setdefault("entries", {})
entry   = entries.setdefault("m365-graph", {})
entry["enabled"] = True

os.makedirs(os.path.dirname(os.path.abspath(config_file)), exist_ok=True)
with open(config_file, "w") as f:
    json.dump(config, f, indent=2)
    f.write("\n")

print(f"Skill 'm365-graph' registered in {config_file}")
PYEOF

echo ""
echo "=== Next steps ==="
echo "1. Follow Microsoft-ENTRA-ID-installation.md to register an app with"
echo "   delegated permissions and enable Public client / device-code flow."
echo "2. Grant the Mail.Send application permission and admin-consent it"
echo "   in Entra ID → App registrations → API permissions."
echo "3. Fill ${INSTALL_DIR}/.env with:"
echo "     TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAIL_SENDER_UPN"
echo "4. Run the first-time login (device-code flow):"
echo "      m365 auth-login"
echo "   Open https://microsoft.com/devicelogin on any device and enter"
echo "   the code shown in the terminal."
echo "5. After login, the skill runs fully headless using cached refresh tokens."
echo "   Mail sending uses the application (client-credentials) flow automatically."
echo "6. Test: m365 calendar-list"
echo ""
echo "Multi-agent note:"
echo "  The skill is shared and visible to all OpenClaw agents on this machine."
echo "  To restrict it to specific agents, set agents.list[].skills in"
echo "  ${OPENCLAW_CONFIG}"
echo ""
echo "Installation complete."
