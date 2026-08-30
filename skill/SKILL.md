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

A CLI (`outlook-scan`, expected on PATH) that queries the mailbox live -
read-only unless write mode is enabled (drafts only). Nothing is stored
locally unless `--save` is passed.

## Ground rules

- **Stdout first.** Print results into the session; use `--save DIR` only
  when the user explicitly wants files kept (extraction, archiving,
  attachments). Confirm or choose an obvious target directory.
- Message ids start with `AAMk`, conversation ids with `AAQk`. Both are
  long — always quote them in commands. `thread` accepts either kind.
- The tool can never send, move, delete, or mark mail. With write mode
  on it can only create drafts for the user to review (see below).

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

## Writing drafts (opt-in, only if write mode is on)

If this install has write mode enabled (`outlook-scan login --write`), you
can create DRAFTS - new messages, in-thread replies and forwards. They
are saved to Drafts and NEVER sent - the user reviews and sends from
Outlook.

```bash
# new message
outlook-scan draft --to "a@x.com" --cc "b@x.com" --subject "..." --body "text"
echo "longer text" | outlook-scan draft --to "a@x.com" --subject "..."

# in-thread reply
outlook-scan reply [--all] --body "text"            "MESSAGE-ID"
echo "longer reply text" | outlook-scan reply       "MESSAGE-ID"
outlook-scan reply --all --body-file /path/reply.txt "MESSAGE-ID"

# forward (attachments of the original are carried over automatically)
outlook-scan forward --to "a@x.com" --body "FYI" "MESSAGE-ID"

# force a real text/plain reply or forward, quoting the original with "> "
outlook-scan reply --plain --body "text" "MESSAGE-ID"
```

`--to/--cc/--bcc` repeat or take comma-separated lists. For bodies longer
than a line or two, prefer piping/`--body-file` over `--body` to avoid
shell quoting trouble.

Text is plain by default; line breaks are preserved. `draft` produces a
true text/plain message. Exchange always builds reply/forward bodies as
HTML - pass `--plain` to override that and get a real text/plain draft
with "> " quoting (mutually exclusive with `--html`). Add `--html` (all
three commands) when the draft should carry formatting - then the text
must be HTML, e.g. `--html --body "<p>Hi <b>there</b></p>"`. Do not mix:
plain text passed with `--html` loses its line breaks, and HTML passed
without `--html` shows up as visible tags.

Typical flows: "draft a reply to X saying Y" → read the mail with
`get`/`thread` for context, compose, then `reply`. "Write a mail to Z
about Y" → `draft`. "Forward that to Z" → `forward`. Show the user the
text you intend to put in the draft before creating it when the wording
matters.

Message ids change when a message is moved between folders. If a write
command returns 404 ErrorItemNotFound, re-run `list`/`search` to get the
current id rather than reusing one from earlier in the session.

Tell the user the draft is in Drafts for review - never imply it was sent.
If it errors "write mode is off", tell them to run `outlook-scan login
--write` once (a one-time consent). Read-only installs simply don't have
this ability - that's expected.

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
