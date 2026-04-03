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
   - `Tasks.ReadWrite.All`

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
m365 calendar-list
```
→ After this step, the token is cached and auto-refreshes silently forever.

### Common issues & fixes

- **"AADSTS65001: The user or administrator has not consented to use the application"**  
  → Admin consent was not granted.  
  Go back to step 2.4 and click **Grant admin consent for [your organization]**.

- **"Microsoft To Do access denied (HTTP 401 Unauthorized)"** when running `todo-*` commands  
  → The application permission `Tasks.ReadWrite.All` is either missing or admin consent has not been granted.  
  Go to Entra ID → App registrations → API permissions, add **Application permission** `Tasks.ReadWrite.All`, then click **Grant admin consent for [your organization]**.

- **"Insufficient privileges"** when accessing mail, calendar or files  
  → Check that the **application permissions** (`Mail.ReadWrite.All`, `Mail.Send`, `Calendars.ReadWrite.All`, `Files.ReadWrite.All`, etc.) are correctly added **and** that admin consent was granted.

- **OneNote not working**  
  → Make sure the **delegated permissions** `Notes.ReadWrite.All` and `offline_access` are added (not application permissions).  
  These are required for OneNote access in most scenarios since early 2025.

- **Secret expired**  
  → Create a new client secret in the Entra portal (Certificates & secrets → New client secret).  
  Update the `CLIENT_SECRET` value in `~/.openclaw/skills/m365-graph/.env`, then run `m365 calendar-list` again to re-authenticate.

You are now fully set up for **unattended Microsoft 365 access** from your OpenClaw agents!  
All subsequent operations (sending mail, updating Excel, creating OneNote pages, etc.) work **without any further login**.

---

## Step 5 – Exchange Online Access (mandatory for Mail access)

Microsoft Exchange Online requires explicit access configuration before the Entra app can read or send mail via the Graph API.
This is separate from the Graph API permissions granted in Step 2.

→ For the full setup guide (RBAC for Applications – recommended) see:
[Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md)
