## Step-by-Step: Entra ID (Microsoft Entra ID) Setup – One-time Configuration

This guide explains how to register a daemon application in Microsoft Entra ID (formerly Azure AD) so your OpenClaw M365 skill can access Microsoft Graph API services **unattended** (without user login after initial setup).

**Minimum requirements**  
- A Microsoft 365 tenant with at least **Microsoft 365 Business Basic** (or higher) license  
- You need **Global Administrator** or **Application Administrator** rights in the tenant  
- Personal/family Microsoft accounts do **not** support daemon / application permissions

### 1. Register a new application

1. Open the Entra admin center:  
   https://entra.microsoft.com

2. Navigate to:  
   **App registrations** → **New registration**

3. Fill in the form:
   - **Name**: `OpenClaw-M365-Daemon` (or any descriptive name)  
   - **Supported account types**: **Accounts in this organizational directory only** (Single tenant)  
   - **Redirect URI**: Leave completely blank  
     *(Daemon apps do not use redirect URIs)*

4. Click **Register**

5. On the **Overview** page, copy and save these two values securely:
   - **Application (client) ID** → this is your `CLIENT_ID`  
   - **Directory (tenant) ID** → this is your `TENANT_ID`

### 2. Add Microsoft Graph API permissions

1. In the app registration → left menu: **API permissions**  
   → Click **Add a permission** → **Microsoft Graph**

2. Select the appropriate permission types:

   **Application permissions** (required for unattended/daemon access):
   - `Mail.ReadWrite.All`  
   - `Mail.Send`  
   - `Calendars.ReadWrite.All`  
   - `Files.ReadWrite.All`
   - `Contacts.ReadWrite`

   **Delegated permissions** (needed for OneNote):
   - `Notes.ReadWrite.All`  
   - `offline_access` (important for refresh tokens)

3. Click **Add permissions** at the bottom of the page.

4. Back on the API permissions overview:  
   Click the button **Grant admin consent for [your organization name]**  
   → Confirm the consent dialog

   → All permissions should now show **Granted for [your org]** (green check mark).  
   *This step is mandatory for application permissions.*

### 3. Create a client secret

1. Left menu: **Certificates & secrets** → tab **Client secrets**  
   → Click **New client secret**

2. Fill in:
   - **Description**: `OpenClaw daemon secret 2026–2028` (or similar)  
   - **Expires**: 24 months (maximum – recommended for initial setup)

3. Click **Add**

4. **Immediately copy the Value** column (not the Secret ID!)  
   → This is your `CLIENT_SECRET`  
   *You will never see this value again after leaving the page.*

**Security recommendation**  
For production environments, consider using a **certificate** instead of a client secret (upload a public certificate under **Certificates** tab). This is more secure and avoids secret rotation issues.

### 4. Summary – The three values you need

After completing the steps above, you should have:

```text
TENANT_ID     = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID     = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET = ~abcdefghijklmnopqrstuvwxyz1234567890abcdef...
```

### Next steps

Paste these three values into the .env file created by the installer:
```bash
~/.openclaw/skills/m365-graph/.env
```
Run the CLI once to perform initial authentication:
```bash
m365
```
→ The script will guide you through the first login/consent (may open a browser or show a URL to copy-paste if running headless).
→ After this step, the token is cached and auto-refreshes silently forever.

### Common issues & fixes

- **"AADSTS65001: The user or administrator has not consented to use the application"**  
  → Admin consent was not granted.  
  Go back to step 2.4 and click **Grant admin consent for [your organization]**.

- **"Insufficient privileges"** when accessing mail, calendar or files  
  → Check that the **application permissions** (`Mail.ReadWrite.All`, `Calendars.ReadWrite.All`, `Files.ReadWrite.All`, etc.) are correctly added **and** that admin consent was granted.

- **OneNote not working**  
  → Make sure the **delegated permissions** `Notes.ReadWrite.All` and `offline_access` are added (not application permissions).  
  These are required for OneNote access in most scenarios since early 2025.

- **Secret expired**  
  → Create a new client secret in the Entra portal (Certificates & secrets → New client secret).  
  Update the `CLIENT_SECRET` value in `~/.openclaw/skills/m365-graph/.env`, then run `m365` again to re-authenticate.

You are now fully set up for **unattended Microsoft 365 access** from your OpenClaw agents!  
All subsequent operations (sending mail, updating Excel, creating OneNote pages, etc.) work **without any further login**.

---

## Step 5 – Exchange Online Access Policy (mandatory for Mail access)

Microsoft Exchange Online requires an explicit access policy before an Entra app can read or send mail via the Graph API.  
This is separate from the Graph API permissions you granted in Step 2.

Use the included PowerShell script **`setup-exchange-policy.ps1`** to configure this automatically.

### Prerequisites

- Windows PowerShell 5.1+ **or** PowerShell 7+ (cross-platform, recommended for Mac/Linux)
- An account with **Exchange Administrator** or **Global Administrator** rights
- The `AppId` (CLIENT_ID) and the **service principal Object ID** of your Entra app  
  *(Object ID is found in Entra admin center → Enterprise applications → your app → Overview)*

### Run the script

```powershell
# From the repository root:
.\setup-exchange-policy.ps1
```

The script will interactively prompt for all required values and guide you through each step.  
You can also supply parameters directly to skip prompts:

```powershell
.\setup-exchange-policy.ps1 `
    -AppId             "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ObjectId          "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -GroupEmailAddress "openclaw-m365@yourdomain.com"
```

### What the script does

1. **Installs / updates** the `ExchangeOnlineManagement` PowerShell module.
2. **Connects** to Exchange Online using Device Code flow (browser-independent, works on Mac/Linux).
3. **Creates a mail-enabled security group** if the specified group does not yet exist, and optionally adds mailbox members.
4. **Configures the access policy** using one of two modes (selectable via `-PolicyMode`):

   | Mode | Description |
   |------|-------------|
   | `RBAC` *(default, recommended)* | Uses **RBAC for Applications** – the current Microsoft standard. Creates a `ServicePrincipal`, `ManagementScope`, and `ManagementRoleAssignment`. |
   | `Legacy` | Uses the deprecated `New-ApplicationAccessPolicy`. Kept for backward compatibility – **Microsoft will retire this**. |

5. **Verifies** the configuration by listing the active role assignments (RBAC mode) or running `Test-ApplicationAccessPolicy` (legacy mode).

### RBAC for Applications – manual reference

If you prefer to run the commands manually:

```powershell
# Install module (once)
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber

# Connect (Device Code flow - works on Mac with no browser pop-up issues)
Connect-ExchangeOnline -Device

# 1. Create the service principal in Exchange (links the Entra app to Exchange)
New-ServicePrincipal `
    -AppId       "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ObjectId    "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -DisplayName "OpenClaw M365 App"

# 2. Create a mail-enabled security group (if not already existing)
New-DistributionGroup `
    -Name               "OpenClaw M365 Exchange Access" `
    -Alias              "openclaw-m365" `
    -PrimarySmtpAddress "openclaw-m365@yourdomain.com" `
    -Type               Security

# Add the mailboxes the app should be able to access
Add-DistributionGroupMember -Identity "openclaw-m365@yourdomain.com" -Member "user@yourdomain.com"

# 3. Create a management scope limited to group members
New-ManagementScope `
    -Name                      "OpenClaw-Scope" `
    -RecipientRestrictionFilter "MemberOfGroup -eq 'openclaw-m365@yourdomain.com'"

# 4. Assign the Application Mail.Read role within that scope
New-ManagementRoleAssignment `
    -Name                "OpenClaw-MailRead" `
    -App                 "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -Role                "Application Mail.Read" `
    -CustomResourceScope "OpenClaw-Scope"

# 5. Verify
Get-ManagementRoleAssignment -App "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" | Format-Table Name, Role, CustomResourceScope
```

### Legacy Application Access Policy - manual reference

> ⚠️ **Deprecated** - Microsoft has replaced Application Access Policies with RBAC for Applications.  
> New policies should not be created. Existing ones will be removed/migrated eventually.

```powershell
# Create the policy (legacy - restricts app to group members only)
New-ApplicationAccessPolicy `
    -AppId              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -PolicyScopeGroupId "openclaw-m365@yourdomain.com" `
    -AccessRight        RestrictAccess `
    -Description        "OpenClaw App – access restricted to group members"

# List all policies
Get-ApplicationAccessPolicy

# Test access for a specific mailbox
Test-ApplicationAccessPolicy -Identity "user@yourdomain.com" -AppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
# → AccessCheckResult should be "Granted" for members of the group

# Remove a legacy policy when migrating to RBAC
Remove-ApplicationAccessPolicy -Identity "<policy-identity-guid>"
```

### Troubleshooting

- **Policy not taking effect** – Exchange policy propagation can take up to 60 minutes. Wait and retry.
- **"ServicePrincipal already exists"** – You can safely ignore this; the script will reuse the existing one.
- **`New-ManagementScope` fails with "filter not valid"** – Make sure the group's `PrimarySmtpAddress` is correct and the group is fully provisioned (wait a few minutes after creation).
- **`Application Mail.Read` role not found** – Ensure you have Exchange Online Plan 1 or higher; this role is not available in Exchange Online Kiosk plans.
