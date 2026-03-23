## Step-by-Step: Exchange Online – Application Access Policy Setup

This guide explains how to restrict the OpenClaw M365 daemon application to specific mailboxes in Exchange Online using an **Application Access Policy**. Without this restriction, an app with `Mail.ReadWrite.All` can access every mailbox in your tenant.

> **Note (as of March 2026):** Application Access Policies are a **legacy mechanism**. Microsoft has replaced them with [RBAC for Applications](https://learn.microsoft.com/en-us/exchange/permissions-exo/rbac-for-applications) and strongly recommends migrating existing policies. A migration guide is included at the end of this document.

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

### 2. Create a mail-enabled Security Group (if not already present)

An Application Access Policy requires a **mail-enabled Security Group** whose members represent the mailboxes the app is allowed to access.

```powershell
# Create the group (alias becomes the local part of the group's email address)
New-DistributionGroup `
  -Name "OpenClaw m365 Exchange Access" `
  -Alias "openclaw-m365-exchange-access" `
  -Type Security

# Add allowed mailboxes as members
# The group email is formed as [alias]@yourdomain.com
Add-DistributionGroupMember `
  -Identity "openclaw-m365-exchange-access@yourdomain.com" `
  -Member "user@yourdomain.com"
```

Replace `yourdomain.com` and the user address with your actual values.  
Repeat `Add-DistributionGroupMember` for every mailbox the application should be permitted to access.

---

### 3. Create the Application Access Policy

This is the core command that ties the Entra ID application to the group:

```powershell
New-ApplicationAccessPolicy `
  -AppId "<YOUR_CLIENT_ID>" `
  -PolicyScopeGroupId "openclaw-m365-exchange-access@yourdomain.com" `
  -AccessRight RestrictAccess `
  -Description "OpenClaw App – access restricted to group members only"
```

| Parameter | Description |
|---|---|
| `-AppId` | The **Application (client) ID** from your Entra ID app registration |
| `-PolicyScopeGroupId` | The primary e-mail address of the mail-enabled Security Group |
| `-AccessRight RestrictAccess` | Allows access **only** for mailboxes that are members of the group |
| `-Description` | Free-text label visible in the policy list |

---

### 4. Verify & Test

```powershell
# List all Application Access Policies in the tenant
Get-ApplicationAccessPolicy

# Test whether a specific mailbox is covered by the policy
Test-ApplicationAccessPolicy `
  -Identity "user@yourdomain.com" `
  -AppId "<YOUR_CLIENT_ID>"
```

Expected output for a mailbox that **is** a group member:

```
AccessCheckResult : Granted
```

Expected output for a mailbox that is **not** a member:

```
AccessCheckResult : Denied
```

> It can take up to **30 minutes** for a newly created policy to be enforced across all Exchange Online endpoints.

---

### 5. (Optional) Remove the policy

```powershell
# List policies to find the Identity GUID
Get-ApplicationAccessPolicy

# Remove by Identity
Remove-ApplicationAccessPolicy -Identity "<policy-guid>"
```

---

### Modern alternative: RBAC for Applications (recommended)

Microsoft has deprecated Application Access Policies in favour of **Role-Based Access Control (RBAC) for Applications**. If you are setting up a new environment or migrating an existing one, use the following approach instead:

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

**When to use which approach:**

| | Application Access Policy (legacy) | RBAC for Applications (modern) |
|---|---|---|
| Status | Deprecated – will be removed | Recommended for new setups |
| Scope | Mail only | Mail, Calendar, Contacts, and more |
| Granularity | Group-level | Full management scope flexibility |
| Migration required | Yes, eventually | No |

For more information see the official Microsoft documentation:  
https://learn.microsoft.com/en-us/exchange/permissions-exo/rbac-for-applications

---

### Common issues & fixes

- **Policy test returns `Denied` even though the mailbox is in the group**  
  → Wait up to 30 minutes for replication. Run `Test-ApplicationAccessPolicy` again afterwards.

- **`New-ApplicationAccessPolicy` returns "group not found"**  
  → Make sure the group was created as a **mail-enabled Security Group** (`-Type Security` in `New-DistributionGroup`). Plain distribution groups are not supported.

- **`Connect-ExchangeOnline` fails on macOS / Linux**  
  → Use the `-Device` flag to switch to Device Code Flow (browser-independent).

- **Policy does not appear in `Get-ApplicationAccessPolicy`**  
  → Verify you are connected to the correct tenant (`Get-OrganizationConfig | Select-Object Name`).
