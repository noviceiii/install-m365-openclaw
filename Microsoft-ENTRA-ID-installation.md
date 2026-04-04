## Step-by-Step: Entra ID (Microsoft Entra ID) Setup – Delegated / Device-Code Flow

This guide explains how to register an application in Microsoft Entra ID (formerly Azure AD)
so your OpenClaw M365 skill can access Microsoft Graph API services using **delegated permissions
and device-code flow**.

The user signs in **once** interactively on any device via `https://microsoft.com/devicelogin`.
Afterwards the app stores a **refresh token** and runs fully **headless** (no browser on the
server required).

**Minimum requirements**
- A Microsoft 365 tenant with at least **Microsoft 365 Business Basic** (or higher) license
- You need **Global Administrator** or **Application Administrator** rights in the tenant
- Personal/family Microsoft accounts do **not** support all delegated permissions listed below

---

### 1. Register a new application

1. Open the Entra admin center:
   https://entra.microsoft.com

2. Navigate to:
   **App registrations** → **New registration**

3. Fill in the form:
   - **Name**: `OpenClaw-M365-Delegated` (or any descriptive name)
   - **Supported account types**: **Accounts in this organizational directory only** (Single tenant)
   - **Redirect URI**: Select **Public client/native (mobile & desktop)** and enter:
     ```
     https://login.microsoftonline.com/common/oauth2/nativeclient
     ```

4. Click **Register**

5. On the **Overview** page, copy and save:
   - **Application (client) ID** → this is your `CLIENT_ID`
   - **Directory (tenant) ID** → this is your `TENANT_ID`

---

### 2. Enable Public client (device-code) flow

1. Left menu: **Authentication**

2. Under **Advanced settings**, set **Allow public client flows** to **Yes**

3. Click **Save**

---

### 3. Add Microsoft Graph API permissions (Delegated)

1. Left menu: **API permissions**
   → Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**

2. Add the following delegated permissions:

   | Permission | Purpose |
   |------------|---------|
   | `User.Read` | Read signed-in user profile |
   | `openid` | OpenID sign-in |
   | `profile` | Read profile claims |
   | `offline_access` | Refresh tokens for headless operation |
   | `Mail.ReadWrite` | Read, write, move, delete mail |
   | `Mail.Send` | Send e-mail |
   | `Calendars.ReadWrite` | Read and write calendar events |
   | `Contacts.ReadWrite` | Read and write contacts |
   | `MailboxFolder.ReadWrite` | Manage mail folders |
   | `Tasks.ReadWrite` | Read and write Microsoft To Do tasks |
   | `Files.ReadWrite` | Read and write OneDrive files |
   | `Notes.ReadWrite` | Read and write OneNote pages |
   | `Sites.ReadWrite.All` | Read and write SharePoint sites |
   | `Bookings.Manage.All` | Full access to Bookings |
   | `Bookings.ReadWrite.All` | Read and write Bookings data |
   | `BookingsAppointment.ReadWrite.All` | Read and write Bookings appointments |
   | `Chat.Create` | Create Teams chats |
   | `Chat.ReadWrite` | Read and write Teams chats |
   | `OnlineMeetings.ReadWrite` | Create and manage Teams meetings |

3. Click **Add permissions**

4. Back on the API permissions overview:
   Click **Grant admin consent for [your organization name]** and confirm.

   → All permissions should show **Granted** (green check mark).

---

### 4. Summary – The two values you need

After completing the steps above, you need:

```text
TENANT_ID  = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID  = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

No client secret is required for delegated / device-code flow.

---

### 5. Configure the skill

Paste the values into the `.env` file created by the installer:

```bash
~/.openclaw/skills/m365-graph/.env
```

It should look like:

```env
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
TOKEN_CACHE_PATH=/home/youruser/.openclaw/credentials/m365_token_cache.bin
```

---

### 6. First-time authentication (device-code flow)

Run the following command on the server:

```bash
m365 auth-login
```

The terminal will display a message like:

```
======================================================================
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABCD1234 to authenticate.
======================================================================
```

Open `https://microsoft.com/devicelogin` on **any** device (your laptop, phone, etc.),
enter the code, and complete the sign-in with an account that has the required permissions.

After successful sign-in:
- The access token and refresh token are cached in `TOKEN_CACHE_PATH`
- The app runs **fully headless** from this point on
- MSAL automatically refreshes the access token using the stored refresh token
- Re-authentication is only needed if the refresh token expires (typically after 90 days of inactivity)

---

### 7. Test the installation

```bash
m365 calendar-list
m365 user-read
m365 mail-list 5
```

---

### Common issues & fixes

- **"Allow public client flows" not enabled**
  → Go to Entra ID → App registrations → Authentication → Advanced settings
  and set **Allow public client flows** to **Yes**.

- **"AADSTS65001: The user or administrator has not consented"**
  → Admin consent was not granted. Go back to Step 3 and click **Grant admin consent**.

- **Token expired – re-authentication needed**
  → Run `m365 auth-login` again to start a new device-code flow.

- **"Insufficient privileges" for Bookings or Teams**
  → Make sure the Bookings / Chat / OnlineMeetings permissions are added and admin consent
  is granted. Some permissions require a Microsoft 365 Business or Enterprise license.

- **Sites.ReadWrite.All not working**
  → Make sure the signed-in user has access to at least one SharePoint site.
  This permission requires admin consent in the tenant.

---

## Exchange Online (optional – for advanced mail features)

For basic mail access (`Mail.ReadWrite`, `Mail.Send`) with delegated permissions,
no additional Exchange configuration is required.

If you encounter issues with mail access in more restrictive tenants, consult:
[Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md)
