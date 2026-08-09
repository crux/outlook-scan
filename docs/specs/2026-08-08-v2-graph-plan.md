# outlook-scan v2 — Microsoft Graph plan (supersedes the AX/UI plan)

Personal Go CLI that extracts mailbox content via the Microsoft Graph API
and writes markdown + JSON files for consumption by Claude Code.
Single user, single account, laptop-local. Not a service.

Supersedes `2026-08-08-v1-plan.md` (AX/accessibility UI scraping). The AX
plan remains on file as the documented fallback should API access ever be
revoked, but is not being built.

## Decision record (2026-08-08)

- Local-access routes verified dead earlier the same day: New Outlook
  16.109 writes no readable local store; legacy sqlite empty; AppleScript
  object model empty. UI scraping worked but was judged inelegant/brittle.
- Device-code consent probe against the tenant SUCCEEDED: delegated
  `Mail.Read` granted with no admin-approval wall (user holds tenant
  admin rights anyway). Graph returned all mail folders including
  Archive (~9k items) and Sent Items (~2k).
- Probe borrowed the first-party "Microsoft Graph Command Line Tools"
  client (14d82eec-204b-4c2f-b7e8-296a70dab67e); the resulting token
  carried broad pre-consented tenant scopes (Directory.ReadWrite.All
  etc.). Production tool MUST use a dedicated app registration scoped to
  `Mail.Read` only, so the cached refresh token is least-privilege.

## Architecture

Plain HTTPS client. No cgo, no Accessibility permissions, Outlook app
irrelevant (may be closed). Go stdlib + `golang.org/x/oauth2` at most.

```
cmd/outlook-scan/       CLI (subcommands: login, folders, scan, thread, status)
internal/msauth/        device-code flow, token cache + silent refresh
internal/graph/         thin REST client (retry, 429 backoff, paging, delta)
internal/export/        markdown + JSON writers (layout carried over from v1 plan)
```

### Auth (internal/msauth)

- One-time Entra app registration ("outlook-scan", public client,
  device-code enabled, delegated `Mail.Read` + `offline_access`).
  Three clicks in the portal or scripted via Graph. Client ID goes in a
  local config file — it is not a secret.
- `outlook-scan login`: device-code flow, prints URL + code, caches
  tokens at `~/.outlook-scan/token.json` (0600). All other commands
  refresh silently and instruct to re-login only when refresh fails.

### Data access (internal/graph)

REVISED 2026-08-09 (decision by Dirk): **on-demand model, no local
mirror.** The tool is a query instrument — list, search, fetch — and the
consuming session (Claude Code) decides what to extract and persist.
Offline archiving is explicitly NOT a feature of this scanner; should
bulk archiving ever be wanted, it would be a separate feature backed by
a local database (sqlite + index), not loose markdown files. No delta
sync, no sync state, no dedup store, no initial sync.

- `GET /me/mailFolders` — folder list with ids and counts.
- List: `GET /me/mailFolders/{name}/messages?$select=<metadata>` with
  optional `$filter` (isRead, receivedDateTime) — recent/unread views.
- Bodies: request `Prefer: outlook.body-content-type="text"` for clean
  plain text (keep HTML variant available behind a flag).
- Threads: `GET /me/messages?$filter=conversationId eq '{id}'` returns
  the complete conversation across ALL folders (Inbox, Sent, Archive) —
  the thread requirement solved server-side.
- `$select` only needed fields: subject, from, toRecipients, ccRecipients,
  receivedDateTime, conversationId, isRead, hasAttachments, body,
  internetMessageId.
- Attachments: metadata always; content download behind `--attachments`.
- Throttling: honor `Retry-After` on 429/503; page size ~50.

### Export (internal/export)

Stdout-first: read operations print markdown (or `--json`) directly, so
an LLM session consumes results without file intermediaries. Files are
written only on explicit `--save DIR`: `YYYY-MM-DD-<slug>.md` per
message, one chronological file per thread.

### CLI

- `outlook-scan login` / `status` / `folders`  (done, Phase 1)
- `outlook-scan list [--folder NAME] [--unread] [--since 7d] [--max 25]`
  — metadata table (date, from, subject, id) of recent mail
- `outlook-scan search "query" [--max 25]` — server-side full-text
  search (`GET /me/messages?$search=...`), metadata table
- `outlook-scan get <message-id> [--save DIR] [--json]` — full message
- `outlook-scan thread <message-id> [--save DIR]` — complete
  conversation via conversationId, chronological, across all folders

### Claude Code integration

Skill (SKILL.md): "check my inbox" → `outlook-scan scan` + read new
markdown; "pull up the thread for X" → `outlook-scan thread`.

## Phases

**Phase 1 — auth + skeleton (~0.5 day).**
Dedicated app registration; msauth with device-code + cache + refresh;
`login`, `status`, `folders` working. Milestone: silent re-auth across
runs with the least-privilege app.

**Phase 2 — query commands (~1 day).**
`list`, `search`, `get`, `thread` with markdown-to-stdout rendering and
`--save`. Milestone: full inbox-triage round-trip from the CLI alone.

**Phase 3 — integration (~0.25 day).**
Claude Code skill ("check my inbox" → list; "find/pull up X" →
search/get/thread), README.

Total: ~2 focused days, est. 700–1,000 LOC.

## Risks (all minor)

- Token on disk = mailbox read access: 0600 perms, least-privilege app,
  document how to revoke (delete app registration / revoke sessions).
- HTML noise in bodies: mitigated by text body preference; quoted-reply
  trimming deferred (Claude handles quoted noise well).
- Graph message ids are long opaque strings — fine for machine use;
  `list`/`search` output orders columns so ids stay out of the way.
