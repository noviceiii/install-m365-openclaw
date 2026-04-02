## Step-by-Step: Exchange Online – RBAC for Applications Setup (Recommended)

This guide explains how to configure Exchange Online so the OpenClaw M365 daemon application can access specific mailboxes via the Graph API.

> **Recommended approach (v0.3.0+):** RBAC for Applications is the Microsoft-recommended method for restricting mailbox access. It replaces the legacy Application Access Policy, which is deprecated and will be removed in a future Exchange Online update.

**Minimum requirements**
- A Microsoft 365 tenant with at least **Microsoft 365 Business Basic** (or higher) license
- You need **Exchange Administrator** or **Global Administrator** rights in the tenant
- The Entra ID app registration must already exist (see [Microsoft-ENTRA-ID-installation.md](Microsoft-ENTRA-ID-installation.md))

---

### 1. Preparation & Connection (one-time)

Install or update the Exchange Online Management PowerShell module and connect to your tenant:

```powershell
# Install/update the module
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber

# Connect using Device Code Flow (recommended on macOS/Linux where browser pop-ups are unreliable)
Connect-ExchangeOnline -Device
```

→ PowerShell prints a device code and a URL.  
→ Open the URL in any browser, enter the code, and sign in with your **admin account** (including MFA).

---

### 2. RBAC for Applications (recommended)

**RBAC for Applications** is the modern, Microsoft-recommended replacement for the deprecated Application Access Policy. Use the included `setup-exchange-policy.ps1` script or follow the manual steps below.

#### Option A: Use the included PowerShell script

```powershell
.\setup-exchange-policy.ps1
```

The script will prompt for your `CLIENT_ID`, `ENTRA_OBJECT_ID`, and the allowed user email, then configure RBAC automatically.

#### Option B: Manual RBAC setup

```powershell
# Step 1 – Create a Service Principal linked to your Entra ID app registration
New-ServicePrincipal `
  -AppId "<YOUR_CLIENT_ID>" `
  -ObjectId "<YOUR_ENTRA_OBJECT_ID>"  # Enterprise Application Object ID from Entra ID

# Step 2 – Create a Management Scope that limits access to specific mailboxes or groups
New-ManagementScope `
  -Name "OpenClaw-MailScope" `
  -RecipientRestrictionFilter "MemberOfGroup -eq 'CN=OpenClaw m365 Exchange Access,...'"

# Step 3 – Assign the appropriate role to the application within that scope
New-ManagementRoleAssignment `
  -Name "OpenClaw-MailRole" `
  -App "<YOUR_CLIENT_ID>" `
  -Role "Application Mail.ReadWrite" `
  -CustomResourceScope "OpenClaw-MailScope"
```

> Replace the filter value in `New-ManagementScope` with the full distinguished name (DN) of your Security Group. You can retrieve it with:
> ```powershell
> Get-Group -Identity "OpenClaw m365 Exchange Access" | Select-Object DistinguishedName
> ```

For more information see the official Microsoft documentation:  
https://learn.microsoft.com/en-us/exchange/permissions-exo/rbac-for-applications

---

### 3. Create a mail-enabled Security Group (if not already present)

RBAC for Applications uses a **mail-enabled Security Group** whose members represent the mailboxes the app is allowed to access.

```powershell
# Create the group (alias becomes the local part of the group's email address)
New-DistributionGroup `
  -Name "OpenClaw m365 Exchange Access" `
  -Alias "openclaw-m365-exchange-access" `
  -Type Security

# Add allowed mailboxes as members
Add-DistributionGroupMember `
  -Identity "openclaw-m365-exchange-access@yourdomain.com" `
  -Member "user@yourdomain.com"
```

Replace `yourdomain.com` and the user address with your actual values.  
Repeat `Add-DistributionGroupMember` for every mailbox the application should be permitted to access.

---

### 4. Verify & Test

```powershell
# List all Management Role Assignments for the application
Get-ManagementRoleAssignment -App "<YOUR_CLIENT_ID>"

# Test mail access
m365 mail-list
```

> It can take up to **30 minutes** for newly configured RBAC permissions to propagate across all Exchange Online endpoints.

---

### Deprecated: Legacy Application Access Policy

> ⚠️ **The Application Access Policy (`New-ApplicationAccessPolicy`) is deprecated** as of March 2026 and will be removed in a future Exchange Online update. Microsoft strongly recommends migrating to RBAC for Applications (see above).

If you are maintaining an existing environment that still uses the legacy policy, you can view and remove it with:

```powershell
# List existing Application Access Policies
Get-ApplicationAccessPolicy

# Remove by Identity
Remove-ApplicationAccessPolicy -Identity "<policy-guid>"
```

After removal, configure RBAC for Applications as described in Section 2 above.

---

### Common issues & fixes

- **Mail access denied (HTTP 403)** after configuring RBAC  
  → Wait up to 30 minutes for permission propagation. Then test with `m365 mail-list`.

- **`New-ServicePrincipal` returns "already exists"**  
  → The service principal was already created. Proceed to Step 2 (Management Scope).

- **`Connect-ExchangeOnline` fails on macOS / Linux**  
  → Use the `-Device` flag to switch to Device Code Flow (browser-independent).

- **Policy does not appear in `Get-ManagementRoleAssignment`**  
  → Verify you are connected to the correct tenant (`Get-OrganizationConfig | Select-Object Name`).
