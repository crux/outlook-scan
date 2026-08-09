// Package smime unwraps opaque S/MIME messages (application/pkcs7-mime).
// Graph delivers such mail as an empty body plus one smime.p7m attachment;
// for signedData the actual message sits inside the CMS structure and can be
// extracted client-side (without verifying the signature). envelopedData is
// genuinely encrypted and cannot be read via the API.
package smime

import (
	"bytes"
	"encoding/asn1"
	"encoding/base64"
	"fmt"
	"io"
	"mime"
	"mime/multipart"
	"mime/quotedprintable"
	"net/mail"
	"strings"

	"github.com/smallstep/pkcs7"
)

// Part is a file extracted from the inner MIME message.
type Part struct {
	Name        string
	ContentType string
	Data        []byte
}

// Result of unwrapping a pkcs7-mime blob.
type Result struct {
	Encrypted bool   // envelopedData: content not readable via API
	Body      string // extracted body text (plain preferred, HTML fallback)
	Parts     []Part // inner attachments
}

var (
	oidEnvelopedData     = asn1.ObjectIdentifier{1, 2, 840, 113549, 1, 7, 3}
	oidAuthEnvelopedData = asn1.ObjectIdentifier{1, 2, 840, 113549, 1, 9, 16, 1, 23}
)

// IsPKCS7Attachment reports whether an attachment looks like an opaque
// S/MIME container.
func IsPKCS7Attachment(name, contentType string) bool {
	return strings.HasSuffix(strings.ToLower(name), ".p7m") ||
		strings.Contains(strings.ToLower(contentType), "pkcs7-mime")
}

// Unwrap parses a CMS blob. Encrypted content is reported, signed content is
// unwrapped into body text and inner attachments. The signature is NOT
// verified - callers must present the result as unverified.
func Unwrap(blob []byte) (*Result, error) {
	if oid, ok := contentTypeOID(blob); ok {
		if oid.Equal(oidEnvelopedData) || oid.Equal(oidAuthEnvelopedData) {
			return &Result{Encrypted: true}, nil
		}
	}
	p7, err := pkcs7.Parse(blob) // handles BER and DER
	if err != nil {
		return nil, fmt.Errorf("parse CMS structure: %w", err)
	}
	if len(p7.Content) == 0 {
		// No extractable content: treat like encrypted rather than failing.
		return &Result{Encrypted: true}, nil
	}
	body, parts := parseInnerMIME(p7.Content)
	return &Result{Body: body, Parts: parts}, nil
}

// contentTypeOID reads the outer ContentInfo type. Best-effort: BER with
// indefinite lengths may not parse here, in which case the pkcs7 library
// (which normalizes BER) decides.
func contentTypeOID(blob []byte) (asn1.ObjectIdentifier, bool) {
	var ci struct {
		ContentType asn1.ObjectIdentifier
		Content     asn1.RawValue `asn1:"optional,explicit,tag:0"`
	}
	if _, err := asn1.Unmarshal(blob, &ci); err != nil {
		return nil, false
	}
	return ci.ContentType, true
}

type mimeState struct {
	plain string
	html  string
	parts []Part
}

// parseInnerMIME walks the extracted RFC 822 message. Best-effort by design:
// whatever was recognized is returned even if some part failed to parse.
func parseInnerMIME(raw []byte) (string, []Part) {
	msg, err := mail.ReadMessage(bytes.NewReader(raw))
	if err != nil {
		// Not a full message - present the payload as-is.
		return strings.TrimSpace(string(raw)), nil
	}
	st := &mimeState{}
	walkPart(func(k string) string { return msg.Header.Get(k) }, msg.Body, st, 0)
	body := st.plain
	if strings.TrimSpace(body) == "" {
		body = st.html
	}
	return strings.TrimSpace(strings.ReplaceAll(body, "\r\n", "\n")), st.parts
}

func walkPart(header func(string) string, r io.Reader, st *mimeState, depth int) {
	if depth > 10 {
		return
	}
	mediaType, params, err := mime.ParseMediaType(header("Content-Type"))
	if err != nil {
		mediaType, params = "text/plain", map[string]string{}
	}

	if strings.HasPrefix(mediaType, "multipart/") {
		mr := multipart.NewReader(r, params["boundary"])
		for {
			p, err := mr.NextPart()
			if err != nil {
				return // io.EOF or malformed part: stop, keep what we have
			}
			walkPart(p.Header.Get, p, st, depth+1)
		}
	}

	data, err := io.ReadAll(transferDecoder(header("Content-Transfer-Encoding"), r))
	if err != nil {
		return
	}

	disposition, dparams, _ := mime.ParseMediaType(header("Content-Disposition"))
	filename := dparams["filename"]
	if filename == "" {
		filename = params["name"]
	}
	if filename != "" || disposition == "attachment" {
		if filename == "" {
			filename = "attachment"
		}
		st.parts = append(st.parts, Part{
			Name:        decodeWord(filename),
			ContentType: mediaType,
			Data:        data,
		})
		return
	}

	text := decodeCharset(data, params["charset"])
	switch mediaType {
	case "text/plain":
		if st.plain == "" {
			st.plain = text
		}
	case "text/html":
		if st.html == "" {
			st.html = text
		}
	}
}

func transferDecoder(cte string, r io.Reader) io.Reader {
	switch strings.ToLower(strings.TrimSpace(cte)) {
	case "base64":
		return base64.NewDecoder(base64.StdEncoding, newBase64Cleaner(r))
	case "quoted-printable":
		return quotedprintable.NewReader(r)
	default:
		return r
	}
}

// decodeCharset converts common legacy charsets to UTF-8 (best-effort;
// ISO-8859-1/Windows-1252 approximated as Latin-1, everything else assumed
// UTF-8-compatible).
func decodeCharset(data []byte, charset string) string {
	switch strings.ToLower(strings.TrimSpace(charset)) {
	case "iso-8859-1", "iso-8859-15", "windows-1252", "latin1":
		var b strings.Builder
		for _, c := range data {
			b.WriteRune(rune(c))
		}
		return b.String()
	default:
		return string(data)
	}
}

// decodeWord decodes RFC 2047 encoded-words in filenames, best-effort.
func decodeWord(s string) string {
	dec := new(mime.WordDecoder)
	if out, err := dec.DecodeHeader(s); err == nil {
		return out
	}
	return s
}

// newBase64Cleaner strips whitespace (CRLF line wrapping) from a base64
// stream so the stdlib decoder accepts it.
func newBase64Cleaner(r io.Reader) io.Reader {
	return &base64Cleaner{r: r}
}

type base64Cleaner struct {
	r   io.Reader
	buf [4096]byte
}

func (c *base64Cleaner) Read(p []byte) (int, error) {
	for {
		max := len(p)
		if max > len(c.buf) {
			max = len(c.buf)
		}
		n, err := c.r.Read(c.buf[:max])
		out := 0
		for _, b := range c.buf[:n] {
			switch b {
			case '\r', '\n', ' ', '\t':
			default:
				p[out] = b
				out++
			}
		}
		if out > 0 || err != nil {
			return out, err
		}
	}
}
