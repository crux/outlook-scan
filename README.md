# outlook-scan

Let Claude Code (or any LLM CLI assistant) read your Outlook / Microsoft
365 inbox: on-demand, read-only mailbox access from the command line —
say "check my inbox", "find the mails about X", "pull up that thread".
Stdout-first markdown, backed by the Microsoft Graph API; the Outlook
client does not need to be installed or running, and no IT tickets or
admin approval are required in most tenants (`outlook-scan setup`
registers its own least-privilege app).

Deliberately **not** a sync/archive tool: nothing is stored locally
unless explicitly requested with `--save`. The consuming session decides
what to extract and keep.

## Commands

```
outlook-scan setup                            one-time setup: create the Entra app and sign in
outlook-scan login                            device-code sign-in (tokens cached, ~90-day sessions)
outlook-scan status                           config, session, signed-in account
outlook-scan folders                          folder list with item/unread counts
outlook-scan list [--folder NAME] [--unread] [--since 7d] [--max N]
outlook-scan search [--max N] "query"         server-side full-text search, all folders
outlook-scan get [--save DIR [--attachments]] MESSAGE-ID
outlook-scan attachments [--save DIR] MESSAGE-ID
outlook-scan thread [--save DIR] MESSAGE-ID|CONVERSATION-ID
outlook-scan reply [--all] [--body TEXT|--body-file FILE] MESSAGE-ID   (write mode only; see below)
```

### Read-only by default; opt-in draft writing

The tool is read-only out of the box. To let it create **reply drafts**
(in-thread, saved to Drafts, **never sent** - you review and send from
Outlook), enable write mode once:

```
outlook-scan login --write      # requests Mail.ReadWrite; --read-only reverts
```

Trade-off: Microsoft has no drafts-only scope, so write mode grants
`Mail.ReadWrite` (full mailbox read/write). It is opt-in per install;
read-only installs keep `Mail.Read` and cannot write. Prefer per-user
consent for write, keeping any org-wide admin consent at `Mail.Read`.

All read commands accept `--json`. Folder names may be display names
(any language) or common aliases (`inbox`, `archive`/`archiv`, `sent`,
`trash`, `junk`, ...).

## Setup

```
go install github.com/crux/outlook-scan/cmd/outlook-scan@latest
outlook-scan setup
```

`setup` explains what it will do, asks once, and handles the rest: one
browser sign-in, a read-only app registration named `outlook-scan` in
your Microsoft 365 tenant, sign-in to it. The temporary setup
permissions it borrows are removed again automatically; the only things
that persist are the app registration, `~/.outlook-scan/config.json`,
and your mail session (`~/.outlook-scan/token.json`, 0600, readable
mail only). `outlook-scan setup --help` shows the knobs.

### Manual registration (locked-down tenants, or by preference)

If `setup` hits a consent wall — or you'd rather click yourself — in
the [Entra admin center](https://entra.microsoft.com) → *App
registrations* → *New registration*:

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

Then hand the two ids from the registration's overview page to the CLI:

```
outlook-scan setup --client-id <application id> --tenant-id <directory id>
```

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

- Pure Go; the single external dependency is `smallstep/pkcs7`, used to
  unwrap opaque S/MIME mail.
- Opaque S/MIME signed messages (empty body + `smime.p7m`) are unwrapped
  automatically: body and real attachments are extracted client-side and
  tagged as unverified. Encrypted (enveloped) mail is marked with a 🔒 -
  its content is not readable via the API by design.
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
