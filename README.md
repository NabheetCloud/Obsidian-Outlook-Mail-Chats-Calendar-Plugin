# Outlook mail, calendar and Teams messages  — Obsidian plugin

Sync your Microsoft 365 into your Obsidian vault — **mail**, **calendar**, and
**Teams messages** — from a single Azure app. Incremental sync, PKCE auth (no
client secret stored), read-only scopes, configurable folder mappings.

> 📖 **Setting up?** See **[SETUP.md](SETUP.md)** for the full step-by-step
> guide: Azure app registration, a per-feature permissions matrix, admin
> consent, and troubleshooting.

> Desktop-only (needs Node `http`/`crypto` for the OAuth loopback). Not for
> Obsidian mobile.

---

## What it does

Three features, each independently toggleable, all under one target folder
(default `10-Mailbox`):

**📧 Mail**
- Reads mail from one or more Outlook folders (Inbox by default).
- Writes **one note per email** to `<target>/<subfolder>/<date> <subject> <id>.md`
  with YAML frontmatter (from/to/cc, dates, conversation id, web link, flags).
- Maintains `_Thread Index.md` per subfolder, grouping messages by conversation.
- Uses **Graph delta queries** so each sync pulls only changes, not the whole
  mailbox. Optional attachment download per folder.

**📅 Calendar** (`10-Mailbox/Calendar/`)
- **One note per event** plus a **sidebar pane** of upcoming meetings.
- Rolling **days-ahead window** (default 14) — upcoming only, no past archiving.
- Optionally **links meetings to related emails** by shared attendees + subject.

**💬 Teams** (`10-Mailbox/Teams/`)
- **One transcript note per conversation** (chats and/or channels) plus a
  regenerated `_Conversation Index.md`.
- New messages are appended incrementally, bounded by a *last-N-days* window.

---

## One-time setup: register an Azure AD app

> 📖 **New here?** See **[SETUP.md](SETUP.md)** for the full step-by-step guide
> with a per-feature permissions matrix (Mail / Calendar / Teams), admin-consent
> notes, and an error-code troubleshooting table. The summary below is the quick
> version.

You need a free Azure AD **app registration** (public client). No secret is
created or stored — auth uses Authorization Code + PKCE.

1. Go to **[Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)** → **New registration**.
2. **Name:** `Obsidian Outlook Mailbox` (anything).
3. **Supported account types:** *Accounts in this organizational directory only*
   (single-tenant) is fine for a work/school account. Multi-tenant also works.
4. **Redirect URI:** platform **Mobile and desktop applications**, value:
   `http://localhost`
   *(Azure allows `http://localhost` on any port for this platform — the plugin
   picks a free port at login time.)*
5. Click **Register**.
6. On the **Overview** page, copy the **Application (client) ID** and the
   **Directory (tenant) ID**.
7. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions** → add:
   - `Mail.Read`
   - `User.Read`
   - `offline_access`
   - `Calendars.Read` (for calendar sync)
   - For **Teams chats**: `Chat.Read`
   - For **Teams channels**: `Team.ReadBasic.All`, `Channel.ReadBasic.All`,
     `ChannelMessage.Read.All` *(these three typically need admin consent)*
   The plugin only requests the Teams scopes when you enable those features, so
   add them only if you'll use Teams sync.
8. If your tenant requires it, click **Grant admin consent** (or ask your admin).
   Work/school tenants often gate consent — without it, login returns
   `AADSTS65001`/`AADSTS90094`.
9. **Authentication** → under *Advanced settings*, ensure **Allow public client
   flows** = **Yes**.

Then in the plugin settings:
- **Application (client) ID** → the value from step 6.
- **Tenant** → your tenant GUID (from step 6), or `organizations` for any
  work/school account, or `common` for any account type.

---

## Install into your vault

From this repo:

```bash
npm install
npm run build          # produces main.js
```

Copy `main.js`, `manifest.json`, and `styles.css` into your vault:

```
<your-vault>/.obsidian/plugins/outlook-mailbox/
```

Or use the helper (set your vault path):

```bash
VAULT="/path/to/your/vault" npm run install:vault    # see scripts/install.mjs
```

Then in Obsidian: **Settings → Community plugins → Installed plugins** → enable
**Outlook Mailbox**. Reload plugins if it doesn't appear.

---

## Configure & sync

Open **Settings → Outlook Mailbox**:

1. Enter the **client ID** and **tenant**, click **Connect** → a browser opens
   for Microsoft sign-in and consent.
2. Set **Target folder** (default `10-Mailbox`).
3. Under **Folders to sync**, each row maps an Outlook folder to a vault
   subfolder:
   | Outlook path | Vault subfolder |
   |---|---|
   | `inbox` | `Inbox` |
   | `archive` | `Archive` |
   | `Projects/ClientX` | `Clients/ClientX` |

   - Single-segment well-known names (`inbox`, `archive`, `sentitems`,
     `deleteditems`, `junkemail`, `drafts`, …) resolve directly.
   - Anything else is matched by **display name**, `/`-separated for nesting.
4. Set **Auto-sync interval** (minutes; `0` disables). Toggle **Sync on
   startup**.

Sync manually anytime via the **mail ribbon icon** or the **"Outlook Mailbox:
Sync now"** command.

---

## Note layout

```
10-Mailbox/
  Inbox/
    _Thread Index.md
    2026-07-04 Quarterly numbers a1b2c3d4e5.md
    2026-07-03 Re Quarterly numbers f6g7h8i9j0.md
  Archive/
    _Thread Index.md
    ...
```

Each email note:

```markdown
---
source: outlook
message_id: AAMk...
conversation_id: AAQk...
subject: Quarterly numbers
from: alice@contoso.com
from_name: Alice Smith
to:
  - nabheet@infinitelocus.com
received: 2026-07-04T09:12:00Z
folder: Inbox
has_attachments: false
web_link: https://outlook.office365.com/...
---

# Quarterly numbers

**From:** Alice Smith <alice@contoso.com>
**To:** Nabheet Madan <nabheet@infinitelocus.com>
**Date:** 2026-07-04T09:12:00Z
**[Open in Outlook](https://outlook.office365.com/...)**

---

<email body as Markdown>
```

---

## How sync works (delta)

- First sync of a folder does a full enumeration and stores an
  `@odata.deltaLink`.
- Subsequent syncs call that delta link → Graph returns only messages added,
  changed, or removed since. Delta links are persisted per resolved folder id
  in the plugin's `data.json`.
- **Updates** overwrite the existing note (same message id → same file name).
- **Removals** move the note to Obsidian trash and drop it from the thread
  index.
- **Reset sync state** (settings → Maintenance) clears delta tokens + thread
  index so the next sync re-enumerates. Existing notes are left in place.

## Calendar

Enable **Settings → Calendar → Sync calendar**.

- **Layout:** one note per event under `10-Mailbox/Calendar/`, plus an
  **Upcoming meetings** sidebar pane (ribbon icon / command to open).
- **Window:** a rolling **days-ahead** window (default 14). Upcoming events only —
  past events are not archived, and events that fall out of the window are not
  deleted from the vault.
- **Email cross-linking:** with *Link related emails* on, each meeting note links
  to emails that share attendees and a similar subject, and vice-versa.
- Calendar uses Graph `calendarView` (expands recurring events into instances);
  it re-reads the window each sync rather than a delta cursor.

## Teams messages

Enable **Settings → Teams → Sync Teams**, then pick **chats** and/or
**channels**. After toggling, click **Reconnect** so the new Graph scopes are
consented (channel scopes usually need an admin).

- **Layout:** one Markdown note per conversation (a running transcript), under
  `10-Mailbox/Teams/` by default, plus a regenerated `_Conversation Index.md`.
  Chats are titled by topic or participants; channels as `Team / Channel`.
- **Incremental:** new messages are **appended** to the transcript — bodies live
  in the note, so `data.json` stays small (only a cursor + last-written
  timestamp per conversation are persisted).
  - **Channels** use Graph message *delta* (a server change-tracking cursor). An
    expired delta token (404/410) transparently re-enumerates, same as mail.
  - **Chats** cannot use delta — Microsoft Graph does **not** support change
    tracking on chat messages (`/chats/{id}/messages/delta` returns
    *"Change tracking is not supported against microsoft.graph.chatMessage"*).
    Instead, chats are listed newest-first and paged only back to the sync
    window; the per-conversation last-written timestamp handles dedup. Practical
    effect: keep the window reasonable (see below) so chat listing stays cheap.
- **Window:** *Sync messages from the last N days* (default 30) bounds how far
  back the first sync reaches **and** how far chat listing pages each run;
  `0` = all history (heavy for chats — every sync walks the full chat).
- **Channels & replies:** channel *posts* sync via delta; replies are fetched
  for posts seen in that sync. A **new reply to an older post** (older than the
  last synced message) won't be picked up until that post next changes — a known
  limitation of channel reply delta. Chats have no such caveat.
- Message **edits** are not re-written (dedup is by "newer than last written");
  deleted and system (join/leave) messages are skipped.

## Observability

- **Status bar**: `📬 Outlook · <last sync time>` / `syncing…` / `error`.
- **Notices** on completion (counts) and on errors.
- **Debug logging** toggle → verbose `[obs-mail]` console output
  (open devtools with `Ctrl/Cmd+Shift+I`).

---

## Security notes

- Auth is **Authorization Code + PKCE** with a loopback redirect. No client
  secret is used or stored.
- The **refresh token** is stored in the plugin's `data.json` inside your vault,
  **in plain text** (same as most Obsidian plugins). Anyone with read access to
  your vault files can use it. Keep your vault private; don't sync `data.json`
  to untrusted locations. Use **Disconnect** to revoke locally, and revoke the
  app under **My Account → Apps & services** / Azure to invalidate server-side.
- All scopes are **read-only** and requested only for the features you enable:
  `User.Read` + `offline_access` (always), `Mail.Read` (mail), `Calendars.Read`
  (calendar), `Chat.Read` (Teams chats), and `Team.ReadBasic.All` +
  `Channel.ReadBasic.All` + `ChannelMessage.Read.All` (Teams channels). The
  plugin never writes to or deletes anything in Microsoft 365.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `AADSTS65001` / consent required | Grant admin consent for the app permissions (step 8). |
| `AADSTS7000218` / public client | Set **Allow public client flows = Yes** (Authentication blade). |
| Login browser opens but nothing happens | Loopback port blocked by firewall; retry, or check devtools console. |
| `Outlook folder not found` | Use the exact **display name**; `/`-separate nested folders; well-known names are lowercase (`sentitems`, not `Sent Items`). |
| Nothing new syncs | State is delta-based; use **Reset sync state** to force full re-enumeration. |
| Conditional Access blocks device | This plugin uses the interactive auth-code flow, not device code — sign in through the opened browser as normal. |

---

## Development

```bash
npm install
npm run dev        # esbuild watch, rebuilds main.js on change
npm run typecheck  # tsc --noEmit
npm run build      # typecheck + minified production bundle
```

Source layout:
- `src/main.ts` — plugin lifecycle, commands, ribbon, status bar, timer.
- `src/auth.ts` — PKCE loopback OAuth flow + token refresh.
- `src/graph.ts` — Graph client: token lifecycle, folder resolution, delta.
- `src/notes.ts` — HTML→Markdown, note + thread-index writing.
- `src/sync.ts` — per-folder delta orchestration + thread bookkeeping.
- `src/settings.ts` — settings UI.
- `src/types.ts` — shared types + defaults.
