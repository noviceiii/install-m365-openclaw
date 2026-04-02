#!/usr/bin/env bash
# =============================================================================
# OpenClaw M365 Graph Skill Installer v0.3.0 – RBAC-first
# =============================================================================
set -euo pipefail

DEFAULT_OPENCLAW_DIR="${HOME}/.openclaw"
DEFAULT_M365_USER="cleotine@yourdomain.com"

echo "=== OpenClaw M365 Graph Skill Installer v0.3.0 (RBAC) ==="

read -p "OpenClaw base directory [${DEFAULT_OPENCLAW_DIR}]: " OPENCLAW_DIR
OPENCLAW_DIR="${OPENCLAW_DIR:-${DEFAULT_OPENCLAW_DIR}}"

read -p "M365 user email for this skill [${DEFAULT_M365_USER}]: " M365_USER
M365_USER="${M365_USER:-${DEFAULT_M365_USER}}"

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
"${VENV_DIR}/bin/pip" install O365 msal msal_extensions python-dotenv python-docx python-pptx openpyxl requests

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
cat > "${INSTALL_DIR}/.env" << EOF
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-app-client-id-here
CLIENT_SECRET=your-app-client-secret-here
TOKEN_CACHE_PATH=${TOKEN_CACHE}
M365_USER_EMAIL=${M365_USER}
EOF

echo ""
echo "=== Next steps ==="
echo "1. Follow Microsoft-ENTRA-ID-installation.md"
echo "2. Run setup-exchange-policy.ps1 (RBAC mode)"
echo "3. Fill ${INSTALL_DIR}/.env with real credentials"
echo "4. Test: m365 calendar-list"
echo ""
echo "Installation complete."
