## Step-by-Step: Entra ID (Microsoft Entra ID) Setup – Hybrid Auth

This guide explains how to register an application in Microsoft Entra ID (formerly Azure AD)
so your OpenClaw M365 skill can access Microsoft Graph API services using **hybrid
authentication**:

- **Delegated permissions + device-code flow** for Calendar, Contacts, OneDrive, Tasks,
  Teams, Bookings, Sites, and reading/managing mail (`GET /me/…`)
- **Application permission `Mail.Send` + client secret** for sending mail via
  `POST /users/{MAIL_SENDER_UPN}/sendMail`

The user signs in **once** interactively on any device via `https://microsoft.com/devicelogin`.
Afterwards the app stores a **refresh token** and runs fully **headless** (no browser on the
server required). Mail sending uses the confidential-client (client-credentials) flow.

The same app registration is both a **public client** (device-code) and a **confidential
client** (client secret).

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
   | `Mail.Send` | Send e-mail (delegated; kept for completeness) |
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

### 4. Add the Mail.Send application permission (required for mail-send)

Unattended `mail-send` uses a confidential client and
`POST /users/{MAIL_SENDER_UPN}/sendMail`. Delegated `/me/sendMail` is not used
for sending.

1. Left menu: **API permissions**
   → Click **Add a permission** → **Microsoft Graph** → **Application permissions**

2. Add:

   | Permission | Purpose |
   |------------|---------|
   | `Mail.Send` | Send mail as the mailbox in `MAIL_SENDER_UPN` |

3. Click **Add permissions**

4. Click **Grant admin consent for [your organization name]** again and confirm.

---

### 5. Create a client secret

1. Left menu: **Certificates & secrets** → **Client secrets** → **New client secret**

2. Add a description and expiry, then click **Add**

3. Copy the **Value** immediately (it is shown only once) → this is your `CLIENT_SECRET`

---

### 6. Summary – The values you need

After completing the steps above, you need:

```text
TENANT_ID       = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID       = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET   = your-client-secret
MAIL_SENDER_UPN = sender@yourdomain.com
```

`MAIL_SENDER_UPN` is the UPN or e-mail of the mailbox the application sends from.

---

### 7. Configure the skill

Paste the values into the `.env` file created by the installer:

```bash
~/.openclaw/skills/m365-graph/.env
```

It should look like:

```env
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=your-client-secret
MAIL_SENDER_UPN=sender@yourdomain.com
TOKEN_CACHE_PATH=/home/youruser/.openclaw/credentials/m365_token_cache.bin
```

---

### 8. Exchange Online RBAC (required before mail-send)

Application-permission `Mail.Send` must be scoped to the sender mailbox with
**RBAC for Applications**. Run this **before** the first `mail-send`:

```powershell
.\setup-exchange-policy.ps1
```

See [Microsoft-Exchange-Policy-installation.md](Microsoft-Exchange-Policy-installation.md)
for the full procedure.

Without this step, Graph may accept the token but Exchange Online will reject
unattended send (typically HTTP 403).

---

### 9. First-time authentication (device-code flow)

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
- The app runs **fully headless** from this point on for delegated features
- MSAL automatically refreshes the access token using the stored refresh token
- Mail sending uses the client secret (application token) and does not use the
  delegated refresh token
- Re-authentication is only needed if the refresh token expires (typically after 90 days of inactivity)

---

### 10. Test the installation

```bash
m365 calendar-list
m365 user-read
m365 mail-list 5
```

`mail-send` additionally requires `CLIENT_SECRET`, `MAIL_SENDER_UPN`, the
application `Mail.Send` permission, and Exchange RBAC from step 8.

---

### Common issues & fixes

- **"Allow public client flows" not enabled**
  → Go to Entra ID → App registrations → Authentication → Advanced settings
  and set **Allow public client flows** to **Yes**.

- **"AADSTS65001: The user or administrator has not consented"**
  → Admin consent was not granted. Go back to Steps 3–4 and click **Grant admin consent**.

- **Token expired – re-authentication needed**
  → Run `m365 auth-login` again to start a new device-code flow.

- **`CLIENT_SECRET is not set` / `MAIL_SENDER_UPN is not set`**
  → Fill both values in `~/.openclaw/skills/m365-graph/.env`. They are required
  for `mail-send` (application / client-credentials flow).

- **Mail send fails with HTTP 403**
  → Confirm the **Mail.Send application** permission is granted and that
  `setup-exchange-policy.ps1` has been run for `MAIL_SENDER_UPN`. Wait up to
  30 minutes for Exchange RBAC to propagate.

- **"Insufficient privileges" for Bookings or Teams**
  → Make sure the Bookings / Chat / OnlineMeetings permissions are added and admin consent
  is granted. Some permissions require a Microsoft 365 Business or Enterprise license.

- **Sites.ReadWrite.All not working**
  → Make sure the signed-in user has access to at least one SharePoint site.
  This permission requires admin consent in the tenant.
