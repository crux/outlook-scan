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
  → `POST /me/messages/{id}/createReply|createReplyAll` with the text as
  `comment` (placed above the quoted original). The draft lands in Drafts;
  the command prints its webLink. Never sends.
- 403 → hint to re-run `login --write` (consent to Mail.ReadWrite).

## Graph / app registration

- Delegated `Mail.ReadWrite` (id 024d486e-b451-40bb-833d-3e66d98c5c73)
  added to the app registration `requiredResourceAccess` so new setups can
  opt in. Existing apps need it added once (portal or PATCH /applications).
- `Mail.ReadWrite` supersets `Mail.Read`; read commands keep working in
  write mode.

## Out of scope

Sending (`Mail.Send`) - deliberately never added, as are delete and move.
Rich/HTML bodies (plain text only; the quoted thread in replies is
preserved by Graph). Attachments on drafts, and forwards
(`createForward`) - both easy later additions if wanted.
