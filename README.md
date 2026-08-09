# outlook-scan

On-demand, read-only access to a Microsoft 365 mailbox from the command
line, designed for consumption by LLM coding assistants (stdout-first,
markdown output). Backed by the Microsoft Graph API — the Outlook client
does not need to be installed or running.

Deliberately **not** a sync/archive tool: nothing is stored locally
unless explicitly requested with `--save`. The consuming session decides
what to extract and keep.

## Commands

```
outlook-scan login                            one-time device-code sign-in (tokens cached, ~90-day sessions)
outlook-scan status                           config, session, signed-in account
outlook-scan folders                          folder list with item/unread counts
outlook-scan list [--folder NAME] [--unread] [--since 7d] [--max N]
outlook-scan search [--max N] "query"         server-side full-text search, all folders
outlook-scan get [--save DIR [--attachments]] MESSAGE-ID
outlook-scan attachments [--save DIR] MESSAGE-ID
outlook-scan thread [--save DIR] MESSAGE-ID|CONVERSATION-ID
```

All read commands accept `--json`. Folder names may be display names
(any language) or common aliases (`inbox`, `archive`/`archiv`, `sent`,
`trash`, `junk`, ...).

## Setup

1. **App registration** (one-time, in your Entra ID tenant — the tool
   needs its own OAuth client identity). In the [Entra admin
   center](https://entra.microsoft.com) → *App registrations* → *New
   registration*:
   - Name e.g. `outlook-scan`; single tenant; no redirect URI needed.
   - *Authentication* → enable **Allow public client flows** (this
     enables the device-code login). No client secret — a CLI can't
     keep one, and the client id is not confidential.
   - *API permissions* → *Microsoft Graph* → *Delegated* → add
     `Mail.Read` (`User.Read` and `offline_access` are granted at
     sign-in automatically).

   You (or each user) consent at first login. If your tenant has user
   consent disabled, an admin must approve the `Mail.Read` grant once —
   it is read-only and limited to the signing user's own mailbox.

2. **Config**: write `~/.outlook-scan/config.json` with the ids shown
   on the registration's overview page:

   ```json
   {"client_id": "<application (client) id>", "tenant_id": "<directory (tenant) id>"}
   ```

3. **Install** (Go 1.22+):

   ```
   go install github.com/crux/outlook-scan/cmd/outlook-scan@latest
   ```

4. **Sign in**: `outlook-scan login` — open the printed URL, enter the
   code, approve. Tokens live in `~/.outlook-scan/token.json` (0600);
   access is limited to reading the signed-in user's own mail.

## Claude Code integration

`skill/SKILL.md` teaches Claude Code the commands and conventions.
Copy it to `~/.claude/skills/outlook-scan/SKILL.md`, personalize the
description with your mailbox address, and "check my inbox" works in
any session.

## Security model

- Delegated `Mail.Read` only: the cached refresh token cannot be
  redeemed for anything beyond reading the signed-in mailbox, because
  the app registration's consent record contains nothing else.
- Revocation: delete `~/.outlook-scan/token.json` locally; revoke the
  user's sign-in sessions or delete the app registration tenant-side.

## Design notes

- Pure Go stdlib; no external dependencies.
- Bodies are requested as plain text (`Prefer:
  outlook.body-content-type="text"`).
- Threads are assembled server-side via `conversationId` across all
  folders and rendered chronologically as a single document.
- Throttling (429/503) is honored via `Retry-After`; complex
  filter/sort combinations fall back to client-side sorting.
- Origin: started as an evaluation of
  [Arkya-AI/outlook-email-scanner](https://github.com/Arkya-AI/outlook-email-scanner)
  (macOS accessibility-tree scraping); rebuilt from scratch on the
  Graph API with an on-demand, no-local-mirror design.
