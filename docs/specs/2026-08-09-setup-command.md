# outlook-scan setup — self-service app registration

Optional convenience command: create the least-privilege Entra app
registration and write `~/.outlook-scan/config.json` without touching
the Entra portal. The portal path (README "Setup") remains the
canonical, always-available route; `setup` is sugar for those who can
and want to use it.

## How the bootstrap works

The chicken-and-egg: creating an app registration requires a Graph
token, but our app doesn't exist yet. Solution: borrow a **first-party
Microsoft public client** for one ephemeral token. Default is
"Microsoft Graph Command Line Tools" (client id
`14d82eec-204b-4c2f-b7e8-296a70dab67e`) — pre-integrated in every
tenant, nothing needs to be installed locally; we only use its client
id in a device-code flow.

Bootstrap steps (all in one command, token held in memory only):

1. Device-code sign-in via the bootstrap client requesting
   `Application.ReadWrite.All` (+ `DelegatedPermissionGrant.ReadWrite.All`
   for the self-clean, see below). The user sees a consent prompt for
   these scopes — this is the "admin-ish" moment; a plain user in a
   locked-down tenant hits the admin-approval wall here and falls back
   to the portal path.
2. Find-or-create the app registration (by display name, default
   `outlook-scan`): public client, single tenant,
   `isFallbackPublicClient`, delegated `Mail.Read` + `User.Read` +
   `offline_access` in `requiredResourceAccess`.
3. Tenant id is read from the bootstrap token's `tid` claim (no extra
   Graph permission needed). Write `config.json` (0600).
4. **Self-clean**: remove `Application.ReadWrite.All` and
   `DelegatedPermissionGrant.ReadWrite.All` from the user's own
   (Principal) consent grant on the bootstrap client — leaving the
   tenant's consent records as they were before setup ran. Skippable
   with `--keep-bootstrap-consent`.
5. Chain into the normal `login` flow against the new app (retry with
   backoff ~60s: fresh registrations can take a moment to propagate).

The bootstrap token is never written to disk and expires within the
hour; the only persisted artifacts are `config.json` and the normal
Mail.Read session from step 5.

## CLI — progressive disclosure

UX principle (decided 2026-08-09): the happy path is ONE command with
ZERO flags; all flexibility exists but only surfaces in `setup --help`.
The main usage text gives setup a single line.

```
outlook-scan setup
    # not yet configured: explain in plain language what will happen
    # (one browser sign-in; creates read-only app "outlook-scan" in
    # your tenant), confirm [Y/n], run the bootstrap, chain into login.
    # already configured: print status + how to redo; change nothing.
```

Hidden depth (help-only flags):
- `--yes`              skip the confirmation (non-interactive use)
- `--name NAME`        app display name (default `outlook-scan`)
- `--via-client ID`    alternative bootstrap public client, for tenants
                       that block the Graph CLI Tools client
- `--keep-bootstrap-consent`  skip the self-clean step
- `--client-id X --tenant-id Y`  manual mode: write config.json after a
                       portal registration, then run login

Output speaks outcomes, not OAuth ("Creating app registration…",
"Cleaning up bootstrap permissions…", "Signed in as …"). The self-clean
is silent default behavior, not a concept the user must evaluate.
README leads with install + `outlook-scan setup`; the portal walkthrough
becomes a "Manual registration" subsection for locked-down tenants.

## Failure modes → messages

- Consent wall at step 1 (`AADSTS65001`/interaction required): explain,
  point to README portal path.
- `403` on app creation: tenant blocks user app registrations → portal
  path or ask an admin.
- App exists but config lost: `--auto` finds it by name and just
  rewrites config (repair case).
- Self-clean failure: warn, print the leftover scope names and the
  portal location to remove them; never fail the whole setup over it.

## Non-goals

- No org-wide ("AllPrincipals") consent is ever created or requested.
- No client secrets, no certificate credentials, no multi-tenant apps.
- Not an admin toolkit: setup manages exactly one app registration and
  only the user's own bootstrap consent, nothing else in the tenant.

## Implementation notes

- New `internal/bootstrap` package (~250 LOC): reuses the msauth
  device-code/token machinery with an override for client id, authority
  (`/organizations`), scopes, and no-persistence mode.
- Graph calls needed: `GET/POST /applications`, `GET /servicePrincipals`,
  `GET/PATCH /oauth2PermissionGrants` — plain GetJSON plus a small POST
  helper on the existing client.
- Prototype: `x.scripts/register_app.py` + `x.scripts/prune_consent.py`
  are the working reference for steps 2 and 4.
