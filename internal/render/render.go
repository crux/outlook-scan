// Package render turns Graph messages into markdown for stdout or files.
package render

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"unicode"

	"github.com/crux/outlook-scan/internal/graph"
)

const timeFmt = "2006-01-02 15:04"

func recipients(rs []graph.Recipient) string {
	parts := make([]string, len(rs))
	for i, r := range rs {
		parts[i] = r.String()
	}
	return strings.Join(parts, ", ")
}

func from(m *graph.Message) string {
	if m.From != nil {
		return m.From.String()
	}
	return "(unknown sender)"
}

// List prints a compact two-line-per-message overview, newest first.
func List(w io.Writer, msgs []graph.Message) {
	if len(msgs) == 0 {
		fmt.Fprintln(w, "no messages")
		return
	}
	for i, m := range msgs {
		mark := " "
		if !m.IsRead {
			mark = "●"
		}
		att := ""
		if m.HasAttachments {
			att = " 📎"
		}
		fmt.Fprintf(w, "%2d. [%s] %s %s — %s%s\n",
			i+1, m.Received.Local().Format(timeFmt), mark, from(&m), m.Subject, att)
		fmt.Fprintf(w, "    id: %s\n", m.ID)
	}
}

// Message renders one full message as markdown.
func Message(m *graph.Message, atts []graph.Attachment) string {
	var b strings.Builder
	fmt.Fprintf(&b, "# %s\n\n", m.Subject)
	fmt.Fprintf(&b, "**From:** %s\n", from(m))
	fmt.Fprintf(&b, "**To:** %s\n", recipients(m.To))
	if len(m.Cc) > 0 {
		fmt.Fprintf(&b, "**Cc:** %s\n", recipients(m.Cc))
	}
	fmt.Fprintf(&b, "**Date:** %s\n", m.Received.Local().Format(timeFmt))
	if !m.IsRead {
		fmt.Fprintf(&b, "**Status:** unread\n")
	}
	if len(atts) > 0 {
		names := make([]string, len(atts))
		for i, a := range atts {
			names[i] = fmt.Sprintf("%s (%s)", a.Name, SizeStr(a.Size))
		}
		fmt.Fprintf(&b, "**Attachments:** %s\n", strings.Join(names, ", "))
	}
	fmt.Fprintf(&b, "**Id:** %s\n", m.ID)
	fmt.Fprintf(&b, "**ConversationId:** %s\n", m.ConversationID)
	b.WriteString("\n---\n\n")
	b.WriteString(body(m))
	b.WriteString("\n")
	return b.String()
}

// Thread renders a whole conversation, oldest first, as one document.
func Thread(msgs []graph.Message) string {
	var b strings.Builder
	subject := ""
	if len(msgs) > 0 {
		subject = msgs[len(msgs)-1].Subject
	}
	fmt.Fprintf(&b, "# Thread: %s (%d messages)\n", subject, len(msgs))
	for _, m := range msgs {
		fmt.Fprintf(&b, "\n## [%s] %s\n\n", m.Received.Local().Format(timeFmt), from(&m))
		fmt.Fprintf(&b, "**To:** %s\n", recipients(m.To))
		if len(m.Cc) > 0 {
			fmt.Fprintf(&b, "**Cc:** %s\n", recipients(m.Cc))
		}
		if m.HasAttachments {
			fmt.Fprintf(&b, "**Attachments:** yes (fetch via `get %s`)\n", m.ID)
		}
		b.WriteString("\n")
		b.WriteString(body(&m))
		b.WriteString("\n")
	}
	return b.String()
}

func body(m *graph.Message) string {
	if m.Body == nil {
		return "(no body)"
	}
	return strings.TrimSpace(strings.ReplaceAll(m.Body.Content, "\r\n", "\n"))
}

// SizeStr formats a byte count human-readably.
func SizeStr(n int) string {
	switch {
	case n >= 1<<20:
		return fmt.Sprintf("%.1f MB", float64(n)/(1<<20))
	case n >= 1<<10:
		return fmt.Sprintf("%.0f KB", float64(n)/(1<<10))
	default:
		return fmt.Sprintf("%d B", n)
	}
}

// SaveRaw writes bytes to dir under the given (sanitized) filename,
// avoiding collisions by appending a counter before the extension.
func SaveRaw(dir, name string, data []byte) (string, error) {
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return "", err
	}
	name = filepath.Base(strings.ReplaceAll(name, "\\", "/")) // strip any path parts
	if name == "" || name == "." || name == "/" {
		name = "attachment"
	}
	ext := filepath.Ext(name)
	stem := strings.TrimSuffix(name, ext)
	path := filepath.Join(dir, name)
	for i := 1; ; i++ {
		if _, err := os.Stat(path); os.IsNotExist(err) {
			break
		}
		path = filepath.Join(dir, fmt.Sprintf("%s-%d%s", stem, i, ext))
	}
	return path, os.WriteFile(path, data, 0o644)
}

// Slug builds a filesystem-safe fragment from a subject line.
func Slug(s string, maxLen int) string {
	var b strings.Builder
	for _, r := range strings.TrimSpace(s) {
		switch {
		case unicode.IsLetter(r) || unicode.IsDigit(r):
			b.WriteRune(unicode.ToLower(r))
		case b.Len() > 0 && !strings.HasSuffix(b.String(), "-"):
			b.WriteRune('-')
		}
	}
	out := strings.Trim(b.String(), "-")
	if len(out) > maxLen {
		out = strings.Trim(out[:maxLen], "-")
	}
	if out == "" {
		out = "no-subject"
	}
	return out
}

// Save writes content to dir with a date-slug name, avoiding collisions.
func Save(dir, date, subject, content string) (string, error) {
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return "", err
	}
	base := date + "-" + Slug(subject, 60)
	path := filepath.Join(dir, base+".md")
	for i := 1; ; i++ {
		if _, err := os.Stat(path); os.IsNotExist(err) {
			break
		}
		path = filepath.Join(dir, fmt.Sprintf("%s-%d.md", base, i))
	}
	return path, os.WriteFile(path, []byte(content), 0o644)
}
