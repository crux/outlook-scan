---
name: outlook-scan
description: Read, search, and fetch the user's Microsoft 365 work email via the outlook-scan CLI. Use when asked to check the inbox, look for/search mails, read a message or a whole thread, or list/download email attachments.
---

<!--
Claude Code skill for outlook-scan. To use it, copy this directory to
~/.claude/skills/outlook-scan/ and personalize the description above
with your mailbox address so Claude knows whose mail it reaches.
-->

# outlook-scan — on-demand mailbox access via Microsoft Graph

A read-only CLI (`outlook-scan`, expected on PATH) that queries the
mailbox live. Nothing is stored locally unless `--save` is passed.

## Ground rules

- **Stdout first.** Print results into the session; use `--save DIR` only
  when the user explicitly wants files kept (extraction, archiving,
  attachments). Confirm or choose an obvious target directory.
- Message ids start with `AAMk`, conversation ids with `AAQk`. Both are
  long — always quote them in commands. `thread` accepts either kind.
- Read-only: the tool cannot send, move, delete, or mark mail.

## Commands

```bash
outlook-scan list [--folder NAME] [--unread] [--since 24h|7d|2w|YYYY-MM-DD] [--max N]
    # metadata, newest first: "N. [date] ● sender — subject 📎" + id line
    # ● = unread, 📎 = has attachments. Default folder Inbox, max 25.
    # Folders: display names or aliases (inbox, archive/archiv, sent,
    # drafts, trash, junk/spam — English and German aliases built in).

outlook-scan search [--max N] "query"
    # server-side full-text search across ALL folders, relevance-ranked

outlook-scan get [--save DIR [--attachments]] "MESSAGE-ID"
    # full message as markdown (headers, attachment names, plain-text body)
    # --save DIR writes .md file; --attachments additionally downloads files

outlook-scan attachments [--save DIR] "MESSAGE-ID"
    # list attachments (name, size, kind); with --save: download them

outlook-scan thread [--save DIR] "MESSAGE-ID-or-CONVERSATION-ID"
    # entire conversation across all folders, chronological, one document

outlook-scan folders | status | login
```

All read commands accept `--json` for structured output.

## Typical flows

- "Check my inbox" → `outlook-scan list --unread` (or `--since 24h`),
  then summarize; offer `get` for anything worth reading in full.
- "Find the mails about X" → `outlook-scan search "X"`, present hits,
  `get`/`thread` on request.
- "Pull up the whole conversation" → `thread` with the message id from a
  previous list/search.
- "Save that mail + attachments" → `get --save DIR --attachments "ID"`.

## S/MIME mail

Opaque-signed mail is unwrapped automatically in get/thread/attachments
- body and real attachments appear normally, tagged "(unwrapped from
opaque S/MIME signature - not verified)". Relay that unverified caveat
when the mail's authenticity matters. Encrypted mail shows "🔒 encrypted
(S/MIME)" - its content is NOT readable via the API; tell the user to
open it in Outlook instead.

## Troubleshooting

- Error `no session — run outlook-scan login` (or refresh failure):
  run `outlook-scan login`, relay the printed URL + code to the user,
  wait for completion. Sessions then last ~90 days silently.
- Folder not found errors list all available folder names.
- Search returns at most 250 results (Graph limit); refine the query.
