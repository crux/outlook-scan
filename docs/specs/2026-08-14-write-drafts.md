# Draft replies (opt-in write mode)

## Goal

Let the tool create **reply drafts** in-thread - "read this mail, draft my
answer" - while keeping the default install strictly read-only. Drafts are
never sent; the user reviews and sends from Outlook.

## Scope trade-off (important)

Microsoft has no "drafts-only" permission. Writing drafts needs
**Mail.ReadWrite** - full mailbox read+write (create/modify/move/delete).
The tool's code only ever calls the reply-draft endpoints, but the token
is capable of more. So write mode is **opt-in per install**:

- Read-only stays the default. `login` requests `Mail.Read`; those tokens
  physically cannot write.
- `login --write` persists `write: true` in config and requests
  `Mail.ReadWrite`. `login --read-only` reverts.
- `reply` refuses unless write mode is on.

Least-privilege is preserved per user by the *requested scope*: a
read-only install's token carries only Mail.Read even if the app is
*capable* of ReadWrite. Caveat: if an admin grants Mail.ReadWrite
**org-wide** (AllPrincipals), any user's refresh token becomes
upgradeable to write - so prefer per-user consent for write users and
keep the org-wide admin consent at Mail.Read only.

## Behavior

- `draft --to ADDR [--cc ADDR] [--bcc ADDR] [--subject TEXT]
  [--body TEXT | --body-file FILE | piped stdin]`
  → `POST /me/messages` with a plain-text body. Address flags repeat or
  take comma-separated lists; at least one `--to` is required and
  addresses are sanity-checked for `@` before the call. Graph does not
  stamp `from` on API-created drafts (Outlook fills it at send time), so
  `get` shows "(unknown sender)" for them - cosmetic, not a defect.
- `reply <id> [--all] [--body TEXT | --body-file FILE | piped stdin]`
  → `POST /me/messages/{id}/createReply|createReplyAll` (no `comment`),
  then PATCH the body with our text merged above the quoted original.
  The draft lands in Drafts; the command prints its webLink. Never sends.
- `forward <id> --to ADDR [--cc ADDR] [--bcc ADDR] [--body ...]`
  → `POST /me/messages/{id}/createForward`; cc/bcc ride along in the
  optional `message` object, body merged as above. Attachments of the
  original are carried over by Graph automatically. Never sent.
- `--html` on draft/reply/forward: the text is HTML instead of plain.
  New drafts set `body.contentType: HTML`. For reply/forward the draft is
  created **without** a comment, and its body (returned by the create
  call, so no extra GET) is then merged and PATCHed: our fragment is
  inserted right after `<body...>`, above Graph's quoted original.
  Plain text is HTML-escaped with newlines converted to `<br>` - Graph
  builds reply/forward bodies as HTML, so raw newlines would otherwise
  collapse into a single paragraph (this was a real defect in the first
  reply implementation, which passed the text as `comment`).
- `--plain` on reply/forward: Exchange always returns an HTML body for
  createReply/createForward, so there is no way to ask it for a
  text/plain reply. Instead the draft body is replaced wholesale: fetch
  the original (Prefer text, so Graph does the HTML→text conversion
  server-side, plus local S/MIME unwrapping), normalize it (CRLF→LF, no
  trailing whitespace, at most one blank line in a row), then PATCH
  `contentType: Text` with our text, an attribution line
  ("On <date>, <sender> wrote:") and the original quoted with "> ".
  Forwards get the classic "---------- Forwarded message ----------"
  header block; attachments are untouched by the body PATCH.
  Mutually exclusive with `--html`.
- 403 → hint to re-run `login --write` (consent to Mail.ReadWrite).
- 404 `ErrorItemNotFound` on a write to an existing message usually means
  a stale id: Graph ids change when a message moves between folders.

## Graph / app registration

- Delegated `Mail.ReadWrite` (id 024d486e-b451-40bb-833d-3e66d98c5c73)
  added to the app registration `requiredResourceAccess` so new setups can
  opt in. Existing apps need it added once (portal or PATCH /applications).
- `Mail.ReadWrite` supersets `Mail.Read`; read commands keep working in
  write mode.

## Out of scope

Sending (`Mail.Send`) - deliberately never added, as are delete and move.
Adding attachments to a draft
(`POST /me/messages/{id}/attachments`) - an easy later addition if
wanted.
