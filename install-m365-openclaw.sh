#!/usr/bin/env bash
# =============================================================================
# OpenClaw M365 Graph Skill Installer v0.5.0
# =============================================================================
set -euo pipefail

DEFAULT_OPENCLAW_DIR="${HOME}/.openclaw"

echo "=== OpenClaw M365 Graph Skill Installer v0.5.0 ==="

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

# .env template
# CLIENT_SECRET and MAIL_SENDER_UPN are required for mail sending via
# application permissions (POST /users/{UPN}/sendMail).
cat > "${INSTALL_DIR}/.env" << EOF
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-app-client-id-here
# Required for application-permission mail sending (client-credentials flow):
CLIENT_SECRET=your-client-secret-here
MAIL_SENDER_UPN=sender@yourdomain.com
TOKEN_CACHE_PATH=${TOKEN_CACHE}
EOF

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
echo "Installation complete."
