# Outlook mail, calendar, and Teams messages — Setup Guide

This guide walks you through everything needed to connect the **Outlook Teams and Calendar**
Obsidian plugin to your Microsoft 365 account: registering a free Azure AD app,
granting the right permissions for the features you want (**Mail**, **Calendar**,
**Teams chats**, **Teams channels**), and configuring the plugin.

It takes about 5–10 minutes. You only do the Azure part **once**.

---

## 0. Before you start

| You need | Notes |
|---|---|
| A Microsoft 365 **work/school** or personal account | Work/school tenants may require an admin for some Teams permissions (called out below). |
| Access to the [Azure Portal](https://portal.azure.com) | Free. Registering an app costs nothing. |
| Obsidian 1.5.0+ desktop | The plugin is **desktop-only** (it opens a local loopback port for sign-in). |

**How sign-in works (so nothing surprises you):** the plugin uses OAuth 2.0
**Authorization Code + PKCE** — the same flow used by native desktop apps. There
is **no client secret**. When you click *Connect*, your browser opens for consent
and redirects back to a short-lived `http://localhost:<port>` listener the plugin
runs. Your **refresh token** is stored locally in the vault (see [Security](#7-security)).

---

## 1. Register the Azure AD app

1. Go to **[Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)**.
2. Click **➕ New registration**.
3. **Name:** anything, e.g. `Obsidian Outlook Teams and Calendar`.
4. **Supported account types:**
   - Work/school only → *Accounts in this organizational directory only (single tenant)*.
   - Want it to work on any Microsoft account → *Accounts in any organizational directory and personal Microsoft accounts*.
5. **Redirect URI** — this part matters, pick the right platform:
   - Platform dropdown: **Mobile and desktop applications** ← *not* "Web", *not* "Single-page application (SPA)".
   - Value: `http://localhost`
   - *(The plugin picks a random free port at sign-in time; Azure allows any port under `http://localhost` for this platform, so you don't register a specific one.)*
6. Click **Register**.

> ⚠️ **Do not create a client secret.** PKCE public clients don't use one.
> Skip the *Certificates & secrets* page entirely — creating a secret there is
> wasted effort and an extra credential to leak.

---

## 2. Turn on public client flows

1. In your new app, open **Authentication** (left sidebar).
2. Scroll to **Advanced settings → Allow public client flows**.
3. Set it to **Yes**, then **Save**.

*(Without this, token redemption for the desktop flow can be rejected.)*

---

## 3. Add API permissions

Open **API permissions → ➕ Add a permission → Microsoft Graph → Delegated
permissions**, and add the scopes for the features you want. **Delegated** (not
Application) — the plugin acts as *you*, seeing only what you can see.

### Permissions matrix

| Feature | Delegated scopes to add | Admin consent usually required? |
|---|---|---|
| **Sign-in (always)** | `User.Read`, `offline_access` | No |
| **Mail** (email → notes) | `Mail.Read` | No |
| **Calendar** (upcoming meetings) | `Calendars.Read` | No |
| **Teams — chats** (1:1 + group DMs) | `Chat.Read` | Sometimes |
| **Teams — channels** (team posts + replies) | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All` | **Yes** (in most tenants) |

**Minimum to get started (mail only):** `User.Read`, `offline_access`, `Mail.Read`.

Add only what you'll use — the plugin requests just the scopes for the features
you enable, so an unused scope sitting in the registration is harmless but
pointless.

> 🔒 **Note on `ChannelMessage.Read.All`:** this is a high-privilege scope. Many
> organizations restrict it and require a **tenant admin** to consent. If you're
> not an admin, you'll need one to approve it before channel sync works. Chats
> (`Chat.Read`) are typically fine with just your own consent.

---

## 4. Grant consent

Still on **API permissions**:

1. If you see a **Grant admin consent for &lt;tenant&gt;** button and you're an
   admin (or have one handy), click it. This pre-approves the scopes for your
   account so the first sign-in is smooth.
2. If you're **not** an admin, you can still try — for user-consentable scopes
   (`Mail.Read`, `Calendars.Read`, `Chat.Read`) the consent prompt appears at
   first sign-in. For `ChannelMessage.Read.All` you'll hit
   `AADSTS65001`/`AADSTS90094` until an admin approves.

---

## 5. Copy your IDs

1. Open the app's **Overview** page (left sidebar, top).
2. Copy **Application (client) ID** — this is your **Client ID**.
3. Copy **Directory (tenant) ID** — this is your **Tenant**.

> 📍 The Client ID lives on the **Overview** page — **not** on *Certificates &
> secrets*. If you're staring at the secrets page looking for it, you're on the
> wrong screen.

---

## 6. Configure the plugin

### Install

- **From Community Plugins** (once published): Settings → Community plugins →
  Browse → search "Outlook Teams and Calendar" → Install → Enable.
- **From source:**
  ```bash
  npm install && npm run build
  VAULT="/path/to/your/vault" npm run install:vault
  ```
  Then enable it under Settings → Community plugins.

### Connect

Open **Settings → Outlook Teams and Calendar**:

1. **Application (client) ID** → paste from step 5.
2. **Tenant** → paste your Directory (tenant) ID, or use `organizations` (any
   work/school account) / `common` (any account type).
3. Click **Connect**. Approve in the browser, return to Obsidian. The status
   should read **Connected as you@company.com**.

> If you add or change features (e.g. turn on Teams) **after** connecting, click
> **Reconnect** so the new scopes are consented.

### Per-feature configuration

**Mail**
- **Target folder** — vault root for all notes (default `10-Mailbox`).
- **Folders to sync** — map Outlook folders → vault subfolders. Use a well-known
  name (`inbox`, `archive`, `sentitems`) or a display-name path (`Projects/ClientX`).
- **Sync mail newer than** — optional `YYYY-MM-DD` floor. Empty = all mail.

**Calendar**
- Toggle **Sync calendar** on.
- **Days ahead** (default 14), **Calendar subfolder** (default `Calendar`),
  **Link related emails** (cross-links meetings ↔ emails by attendees/subject).
- **Open sidebar** shows the *Upcoming meetings* pane (also on the ribbon).

**Teams**
- Toggle **Sync Teams** on, then pick **chats** and/or **channels**.
- **Teams subfolder** (default `Teams`) — one transcript note per conversation.
- **Sync messages from the last N days** (default 30) bounds the first sync's
  reach; `0` = all history.
- After toggling Teams, click **Reconnect** (new scopes).

---

## First sync & what to expect

- Click the **ribbon mail icon**, or run the command **Outlook Teams and Calendar: Sync now**,
  or use **Settings → Maintenance → Sync now**.
- Progress shows in the **status bar** (`📬 …: writing 240/1200…`) and in the
  developer console (Ctrl/Cmd-Shift-I) under the `[obs-mail]` prefix.
- The **first** sync is the slow one (full enumeration). After that, everything
  is incremental (delta) and fast.
- **Teams first sync** is the heaviest: each chat pulls history back to your
  *last-N-days* window. If it's too much, lower that number and **Reset sync state**.
- **Stop** anytime via Settings → Maintenance → Stop (or the *Stop sync* command).
  Already-written notes are kept; the next sync resumes where it left off.

---

## Troubleshooting

| Symptom / error | Cause & fix |
|---|---|
| `AADSTS9002326` (cross-origin / SPA) | Your redirect URI platform is **SPA** or **Web**. Delete it and add `http://localhost` under **Mobile and desktop applications**. |
| `AADSTS65001` / `AADSTS90094` | Consent not granted for a scope. Ask an admin to **Grant admin consent** (common for `ChannelMessage.Read.All`). |
| `AADSTS700016` / `invalid_client` | Wrong Client ID, or the app is in a different tenant than the **Tenant** value. Recheck the Overview page. |
| `AADSTS50011` (redirect mismatch) | Redirect URI isn't `http://localhost` on the *Mobile and desktop* platform. |
| Sign-in window never returns | A firewall is blocking the local loopback port, or you closed the browser tab before approving. Try **Connect** again. |
| "Running but nothing in logs" | Open the console; baseline `[obs-mail]` progress prints regardless of the Debug toggle. If truly stuck, any single request throws after 60 s and logs a labelled error. |
| Teams sync silent for a while | Normal on first run while it lists chats and pulls history — watch for `Teams: … chat(s)…` and per-page `Teams: page N …` heartbeats. |
| Teams `…(403)…` in the log | `Chat.Read` (or the channel scopes) wasn't consented. Add the scope, **Reconnect**, and get admin consent if needed. |
| Teams `400 … Change tracking is not supported against microsoft.graph.chatMessage` | Fixed in current builds — chat sync no longer uses delta. Redeploy the latest `main.js` and reload the plugin. |
| Duplicate/renamed notes after an update | If the filename scheme changed, clear the affected subfolder and **Reset sync state** before re-syncing. |

Toggle **Debug logging** (Settings → Sync) for verbose per-item output.

---

## 7. Security

- Your OAuth **refresh token** is stored in plain text in the plugin's
  `data.json` inside your vault. It is **not** encrypted at rest. Anyone with
  read access to your vault files can use it to read your mail/calendar/Teams.
- There is **no client secret** to manage (PKCE).
- All Graph permissions are **read-only** and **delegated** — the plugin never
  writes to or deletes anything in Microsoft 365, and only ever sees what your
  own account can see.
- Revoke access anytime at
  [My Account → Apps & services / permissions](https://myaccount.microsoft.com/),
  or remove the app registration in Azure. You can also disconnect from
  **Settings → Outlook Teams and Calendar → (disconnect icon)**.

---

## Scope reference (what each permission is for)

| Scope | Grants the plugin | Used by |
|---|---|---|
| `User.Read` | Your profile (name/email, for the "Connected as" label) | Sign-in |
| `offline_access` | A refresh token (stay signed in without re-consenting) | Sign-in |
| `Mail.Read` | Read your mail in all folders | Mail |
| `Calendars.Read` | Read your calendars/events | Calendar |
| `Chat.Read` | Read your Teams chat messages | Teams chats |
| `Team.ReadBasic.All` | List teams you're a member of | Teams channels |
| `Channel.ReadBasic.All` | List channels in those teams | Teams channels |
| `ChannelMessage.Read.All` | Read channel messages/replies | Teams channels |
