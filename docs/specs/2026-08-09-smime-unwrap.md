# Opaque S/MIME unwrapping

## Problem

Graph delivers opaque S/MIME messages as an empty body plus a single
`smime.p7m` attachment (`application/pkcs7-mime`). Without handling,
such mail renders content-less and its real attachments are invisible.
Common in German university/clinic environments (HARICA/GEANT PKI),
where outbound mail is routinely opaque-signed.

## Behavior

Detection: body empty (or whitespace) AND an attachment matching
`*.p7m` / content type containing `pkcs7-mime`. Applied in `get`,
`thread`, and `attachments` (there: only when the container is the
sole attachment).

- CMS `signedData` → extract the inner RFC 822 message client-side,
  render body (text/plain preferred, HTML fallback, Latin-1 charsets
  converted) and inner attachments normally. Body is prefixed:
  `(unwrapped from opaque S/MIME signature - not verified)`.
  The signature is NOT verified - the tag must always say so.
- CMS `envelopedData` / `authEnvelopedData` → body reads
  `🔒 encrypted (S/MIME) - content not readable via the API`.
- `--save`/`--attachments` write the unwrapped inner files, never the
  `p7m` container (except for encrypted mail, where the container is
  all there is).
- Unwrap failures warn on stderr and fall back to the verbatim view -
  never fail the command over it.

## Implementation

`internal/smime`: CMS type via ASN.1 ContentInfo OID (best-effort,
BER-tolerant fallback through the parser), `signedData` payload via
github.com/smallstep/pkcs7 (the project's single external dependency),
inner MIME walked with stdlib (net/mail, mime/multipart,
quoted-printable/base64 decoding, RFC 2047 filenames).

Deliberately out of scope for now: signature verification
(`--verify-smime` as a possible later flag), non-Latin-1 legacy
charsets, decryption of enveloped mail (impossible without the
recipient's private key by design).

Verified 2026-08-09 against a real opaque-signed mail (German
university hospital, HARICA chain): body and 16 inner attachments
including the needed form extracted correctly; `file` confirms valid
payloads.
