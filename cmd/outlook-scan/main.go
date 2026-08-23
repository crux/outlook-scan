// outlook-scan is an on-demand mailbox query tool backed by Microsoft
// Graph: list, search, and fetch messages or whole threads as markdown,
// designed for consumption by LLM sessions (stdout-first, --save opt-in).
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/crux/outlook-scan/internal/bootstrap"
	"github.com/crux/outlook-scan/internal/graph"
	"github.com/crux/outlook-scan/internal/msauth"
	"github.com/crux/outlook-scan/internal/render"
	"github.com/crux/outlook-scan/internal/smime"
)

const smimeUnwrapNote = "(unwrapped from opaque S/MIME signature - not verified)"
const smimeEncryptedNote = "🔒 encrypted (S/MIME) - content not readable via the API"

// maybeUnwrapSMIME detects an opaque S/MIME message (empty body + pkcs7
// container attachment), downloads and unwraps it. Returns nil when the
// message is not S/MIME or unwrapping failed (warned on stderr).
func maybeUnwrapSMIME(c *graph.Client, m *graph.Message, atts []graph.Attachment) *smime.Result {
	if m.Body != nil && strings.TrimSpace(m.Body.Content) != "" {
		return nil
	}
	for _, a := range atts {
		if !smime.IsPKCS7Attachment(a.Name, a.ContentType) {
			continue
		}
		blob, err := c.AttachmentContent(m.ID, a.ID)
		if err != nil {
			fmt.Fprintln(os.Stderr, "warning: S/MIME container download failed:", err)
			return nil
		}
		res, err := smime.Unwrap(blob)
		if err != nil {
			fmt.Fprintln(os.Stderr, "warning: S/MIME unwrap failed:", err)
			return nil
		}
		return res
	}
	return nil
}

// applySMIME rewrites message body and display attachments from an unwrap
// result.
func applySMIME(m *graph.Message, res *smime.Result) []graph.Attachment {
	if res.Encrypted {
		m.Body = &graph.ItemBody{ContentType: "text", Content: smimeEncryptedNote}
		return nil
	}
	m.Body = &graph.ItemBody{ContentType: "text", Content: smimeUnwrapNote + "\n\n" + res.Body}
	var atts []graph.Attachment
	for _, p := range res.Parts {
		atts = append(atts, graph.Attachment{
			Type: "#microsoft.graph.fileAttachment",
			Name: p.Name, Size: len(p.Data), ContentType: p.ContentType,
		})
	}
	return atts
}

func usage() {
	fmt.Fprintln(os.Stderr, `usage: outlook-scan <command> [flags] [args]

commands:
  setup                       one-time setup: create the Entra app and sign in
  login       [flags]         sign in via device code (one-time; tokens cached)
  status                      show config, session, and signed-in user
  folders                     list mail folders with item counts
  list        [flags]         recent messages (metadata) from a folder
  search      [flags] <query> server-side full-text search across all folders
  get         [flags] <id>    one full message as markdown
  thread      [flags] <id>    whole conversation (by message or conversation id)
  attachments [flags] <id>    list a message's attachments; --save downloads them
  draft       [flags]         compose a new DRAFT message (never sends; needs write mode)
  reply       [flags] <id>    create an in-thread reply DRAFT (never sends; needs write mode)
  forward     [flags] <id>    create a forward DRAFT, attachments included (never sends)

flags (login):       --write (enable draft writing, Mail.ReadWrite)  --read-only (revert)
flags (list):        --folder NAME  --unread  --since 7d|2w|2026-01-15  --max N  --json
flags (search):      --max N  --json
flags (get):         --save DIR  --attachments (with --save: download files too)  --json
flags (thread):      --save DIR  --json
flags (attachments): --save DIR  --json
flags (draft):       --to ADDR  --cc ADDR  --bcc ADDR  --subject TEXT  --body TEXT | --body-file FILE | (piped stdin)
flags (reply):       --all (reply to all)  --body TEXT | --body-file FILE | (piped stdin)
flags (forward):     --to ADDR  --cc ADDR  --bcc ADDR  --body TEXT | --body-file FILE | (piped stdin)

Flags come before positional arguments.`)
}

func main() {
	if len(os.Args) < 2 {
		usage()
		os.Exit(2)
	}
	cmds := map[string]func([]string) error{
		"setup":       cmdSetup,
		"login":       cmdLogin,
		"status":      func([]string) error { return cmdStatus() },
		"folders":     func([]string) error { return cmdFolders() },
		"list":        cmdList,
		"search":      cmdSearch,
		"get":         cmdGet,
		"thread":      cmdThread,
		"attachments": cmdAttachments,
		"draft":       cmdDraft,
		"reply":       cmdReply,
		"forward":     cmdForward,
	}
	cmd, ok := cmds[os.Args[1]]
	if !ok {
		usage()
		os.Exit(2)
	}
	if err := cmd(os.Args[2:]); err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
}

func client() (*graph.Client, error) {
	auth, err := msauth.Load()
	if err != nil {
		return nil, err
	}
	return graph.New(auth), nil
}

func cmdSetup(args []string) error {
	fs := flag.NewFlagSet("setup", flag.ExitOnError)
	yes := fs.Bool("yes", false, "skip the confirmation prompt")
	name := fs.String("name", "outlook-scan", "app registration display name")
	viaClient := fs.String("via-client", "", "alternative bootstrap public client id")
	keep := fs.Bool("keep-bootstrap-consent", false, "do not remove the setup permissions afterwards")
	clientID := fs.String("client-id", "", "manual mode: client id of a portal-created app registration")
	tenantID := fs.String("tenant-id", "", "manual mode: directory (tenant) id")
	fs.Parse(args)

	// Manual mode: portal registration done by hand, just write config + login.
	if *clientID != "" || *tenantID != "" {
		if *clientID == "" || *tenantID == "" {
			return fmt.Errorf("manual mode needs both --client-id and --tenant-id")
		}
		if err := msauth.WriteConfig(msauth.Config{ClientID: *clientID, TenantID: *tenantID}); err != nil {
			return err
		}
		fmt.Println("Config written.")
		return cmdLogin(nil)
	}

	// Already configured: report and change nothing.
	if auth, err := msauth.Load(); err == nil {
		cfg := auth.Config()
		fmt.Printf("Already set up (tenant %s, app %s).\n", cfg.TenantID, cfg.ClientID)
		if auth.HasSession() {
			fmt.Println("A session exists — you're good. See `outlook-scan status`.")
		} else {
			fmt.Println("No session yet — run `outlook-scan login`.")
		}
		fmt.Println("To redo setup, delete ~/.outlook-scan/config.json first.")
		return nil
	}

	fmt.Println("outlook-scan setup will:")
	fmt.Println("  1. Ask you to sign in once in the browser (device code).")
	fmt.Printf("  2. Create a read-only app registration %q in your Microsoft 365 tenant.\n", *name)
	fmt.Println("  3. Sign you in to it. Mail access is read-only, your own mailbox only.")
	fmt.Println("(Prefer doing this by hand? See the manual registration section in the README.)")
	if !*yes {
		fmt.Print("Proceed? [Y/n] ")
		var answer string
		fmt.Scanln(&answer)
		answer = strings.ToLower(strings.TrimSpace(answer))
		if answer != "" && answer != "y" && answer != "yes" {
			fmt.Println("Aborted — nothing was changed.")
			return nil
		}
	}
	return bootstrap.Run(os.Stdout, bootstrap.Options{
		Name:        *name,
		ViaClient:   *viaClient,
		KeepConsent: *keep,
	})
}

func cmdLogin(args []string) error {
	fs := flag.NewFlagSet("login", flag.ExitOnError)
	write := fs.Bool("write", false, "enable draft writing for this install (requests Mail.ReadWrite)")
	readOnly := fs.Bool("read-only", false, "revert this install to read-only (Mail.Read)")
	fs.Parse(args)
	if *write && *readOnly {
		return fmt.Errorf("--write and --read-only are mutually exclusive")
	}
	auth, err := msauth.Load()
	if err != nil {
		return err
	}
	desiredWrite := auth.WriteMode()
	switch {
	case *write:
		desiredWrite = true
		fmt.Println("The sign-in below asks you to consent to Mail.ReadWrite (draft writing).")
	case *readOnly:
		desiredWrite = false
	}
	return auth.LoginWith(os.Stdout, desiredWrite)
}

// addrList collects an address flag that may repeat and/or hold a
// comma-separated list: --to a@x --to "b@x, c@x".
type addrList []string

func (l *addrList) String() string { return strings.Join(*l, ", ") }

func (l *addrList) Set(v string) error {
	for _, part := range strings.Split(v, ",") {
		if p := strings.TrimSpace(part); p != "" {
			*l = append(*l, p)
		}
	}
	return nil
}

// checkAddrs rejects anything that clearly is not an email address before a
// draft is created from it.
func checkAddrs(lists ...addrList) error {
	for _, l := range lists {
		for _, a := range l {
			if !strings.Contains(a, "@") {
				return fmt.Errorf("%q does not look like an email address", a)
			}
		}
	}
	return nil
}

// writeAuth loads the session and enforces the opt-in write gate.
func writeAuth() (*msauth.Auth, error) {
	auth, err := msauth.Load()
	if err != nil {
		return nil, err
	}
	if !auth.WriteMode() {
		return nil, fmt.Errorf("write mode is off - enable it once with: outlook-scan login --write")
	}
	return auth, nil
}

// draftError adds a consent hint to permission failures.
func draftError(err error) error {
	if strings.Contains(err.Error(), "HTTP 403") {
		return fmt.Errorf("permission denied - re-run `outlook-scan login --write` to consent to Mail.ReadWrite: %w", err)
	}
	return err
}

// reportDraft prints the standard confirmation for a created draft.
func reportDraft(what string, d *graph.DraftRef) {
	fmt.Printf("Draft %s created in your Drafts folder - review and send from Outlook.\n", what)
	if d.WebLink != "" {
		fmt.Println("open:", d.WebLink)
	}
}

func cmdForward(args []string) error {
	fs := flag.NewFlagSet("forward", flag.ExitOnError)
	var to, cc, bcc addrList
	fs.Var(&to, "to", "recipient address (repeatable, or comma-separated)")
	fs.Var(&cc, "cc", "cc address (repeatable, or comma-separated)")
	fs.Var(&bcc, "bcc", "bcc address (repeatable, or comma-separated)")
	body := fs.String("body", "", "comment placed above the forwarded message")
	bodyFile := fs.String("body-file", "", "read the comment from a file")
	fs.Parse(args)
	if fs.NArg() != 1 || len(to) == 0 {
		return fmt.Errorf("usage: outlook-scan forward --to ADDR [--cc ADDR] [--bcc ADDR] [--body TEXT|--body-file FILE] <message-id>")
	}
	if err := checkAddrs(to, cc, bcc); err != nil {
		return err
	}
	auth, err := writeAuth()
	if err != nil {
		return err
	}
	text, err := readBodyInput(*body, *bodyFile)
	if err != nil {
		return err
	}
	d, err := graph.New(auth).CreateForwardDraft(fs.Arg(0), to, cc, bcc, text)
	if err != nil {
		return draftError(err)
	}
	reportDraft("forward to "+to.String(), d)
	return nil
}

func cmdDraft(args []string) error {
	fs := flag.NewFlagSet("draft", flag.ExitOnError)
	var to, cc, bcc addrList
	fs.Var(&to, "to", "recipient address (repeatable, or comma-separated)")
	fs.Var(&cc, "cc", "cc address (repeatable, or comma-separated)")
	fs.Var(&bcc, "bcc", "bcc address (repeatable, or comma-separated)")
	subject := fs.String("subject", "", "subject line")
	body := fs.String("body", "", "message text (inline)")
	bodyFile := fs.String("body-file", "", "read message text from a file")
	fs.Parse(args)

	if len(to) == 0 {
		return fmt.Errorf("usage: outlook-scan draft --to ADDR [--cc ADDR] [--bcc ADDR] --subject TEXT [--body TEXT|--body-file FILE]")
	}
	if err := checkAddrs(to, cc, bcc); err != nil {
		return err
	}
	auth, err := writeAuth()
	if err != nil {
		return err
	}
	text, err := readBodyInput(*body, *bodyFile)
	if err != nil {
		return err
	}

	d, err := graph.New(auth).CreateDraft(graph.NewDraft{
		To: to, Cc: cc, Bcc: bcc, Subject: *subject, Body: text,
	})
	if err != nil {
		return draftError(err)
	}
	reportDraft("to "+to.String(), d)
	return nil
}

func cmdReply(args []string) error {
	fs := flag.NewFlagSet("reply", flag.ExitOnError)
	all := fs.Bool("all", false, "reply to all recipients")
	body := fs.String("body", "", "reply text (inline)")
	bodyFile := fs.String("body-file", "", "read reply text from a file")
	fs.Parse(args)
	if fs.NArg() != 1 {
		return fmt.Errorf("usage: outlook-scan reply [--all] [--body TEXT|--body-file FILE] <message-id>")
	}
	auth, err := writeAuth()
	if err != nil {
		return err
	}
	text, err := readBodyInput(*body, *bodyFile)
	if err != nil {
		return err
	}
	if strings.TrimSpace(text) == "" {
		return fmt.Errorf("empty reply body - provide --body, --body-file, or pipe text via stdin")
	}
	d, err := graph.New(auth).CreateReplyDraft(fs.Arg(0), *all, text)
	if err != nil {
		return draftError(err)
	}
	kind := "reply"
	if *all {
		kind = "reply-all"
	}
	reportDraft(kind, d)
	return nil
}

// readBodyInput resolves the reply text from --body, --body-file, or piped
// stdin (in that order).
func readBodyInput(body, bodyFile string) (string, error) {
	if body != "" {
		return body, nil
	}
	if bodyFile != "" {
		b, err := os.ReadFile(bodyFile)
		return string(b), err
	}
	if stat, _ := os.Stdin.Stat(); stat != nil && stat.Mode()&os.ModeCharDevice == 0 {
		b, err := io.ReadAll(os.Stdin)
		return string(b), err
	}
	return "", nil
}

func cmdStatus() error {
	auth, err := msauth.Load()
	if err != nil {
		return err
	}
	cfg := auth.Config()
	fmt.Printf("tenant:    %s (%s)\n", cfg.TenantName, cfg.TenantID)
	fmt.Printf("client_id: %s\n", cfg.ClientID)
	if !auth.HasSession() {
		fmt.Println("session:   none — run `outlook-scan login`")
		return nil
	}
	if exp := auth.ExpiresAt(); time.Now().Before(exp) {
		fmt.Printf("session:   access token valid for %s\n", time.Until(exp).Round(time.Second))
	} else {
		fmt.Println("session:   access token expired (will refresh silently)")
	}
	me, err := graph.New(auth).Me()
	if err != nil {
		return err
	}
	fmt.Printf("account:   %s (%s)\n", me.UserPrincipalName, me.DisplayName)
	return nil
}

func cmdFolders() error {
	c, err := client()
	if err != nil {
		return err
	}
	folders, err := c.Folders()
	if err != nil {
		return err
	}
	for _, f := range folders {
		fmt.Printf("%-25s %6d items %6d unread\n", f.DisplayName, f.TotalItems, f.UnreadItems)
	}
	return nil
}

// parseSince accepts 7d, 24h, 2w style durations or a YYYY-MM-DD date.
func parseSince(s string) (time.Time, error) {
	if s == "" {
		return time.Time{}, nil
	}
	if t, err := time.ParseInLocation("2006-01-02", s, time.Local); err == nil {
		return t, nil
	}
	if len(s) >= 2 {
		if n, err := strconv.Atoi(s[:len(s)-1]); err == nil && n > 0 {
			switch s[len(s)-1] {
			case 'h':
				return time.Now().Add(-time.Duration(n) * time.Hour), nil
			case 'd':
				return time.Now().AddDate(0, 0, -n), nil
			case 'w':
				return time.Now().AddDate(0, 0, -7*n), nil
			}
		}
	}
	return time.Time{}, fmt.Errorf("cannot parse --since %q (use 24h, 7d, 2w, or YYYY-MM-DD)", s)
}

func printJSON(v any) error {
	enc := json.NewEncoder(os.Stdout)
	enc.SetIndent("", "  ")
	return enc.Encode(v)
}

func cmdList(args []string) error {
	fs := flag.NewFlagSet("list", flag.ExitOnError)
	folder := fs.String("folder", "Inbox", "folder display name or well-known name")
	unread := fs.Bool("unread", false, "unread messages only")
	since := fs.String("since", "", "only messages newer than (24h, 7d, 2w, YYYY-MM-DD)")
	maxN := fs.Int("max", 25, "maximum messages")
	asJSON := fs.Bool("json", false, "raw JSON output")
	fs.Parse(args)

	sinceT, err := parseSince(*since)
	if err != nil {
		return err
	}
	c, err := client()
	if err != nil {
		return err
	}
	msgs, err := c.List(graph.ListOpts{
		Folder: *folder, UnreadOnly: *unread, Since: sinceT, Max: *maxN,
	})
	if err != nil {
		return err
	}
	if *asJSON {
		return printJSON(msgs)
	}
	render.List(os.Stdout, msgs)
	return nil
}

func cmdSearch(args []string) error {
	fs := flag.NewFlagSet("search", flag.ExitOnError)
	maxN := fs.Int("max", 25, "maximum results")
	asJSON := fs.Bool("json", false, "raw JSON output")
	fs.Parse(args)
	query := strings.TrimSpace(strings.Join(fs.Args(), " "))
	if query == "" {
		return fmt.Errorf("usage: outlook-scan search [--max N] <query>")
	}
	c, err := client()
	if err != nil {
		return err
	}
	msgs, err := c.Search(query, *maxN)
	if err != nil {
		return err
	}
	if *asJSON {
		return printJSON(msgs)
	}
	render.List(os.Stdout, msgs)
	return nil
}

func cmdGet(args []string) error {
	fs := flag.NewFlagSet("get", flag.ExitOnError)
	save := fs.String("save", "", "write markdown to this directory instead of stdout")
	withAtts := fs.Bool("attachments", false, "with --save: also download file attachments")
	asJSON := fs.Bool("json", false, "raw JSON output")
	fs.Parse(args)
	if fs.NArg() != 1 {
		return fmt.Errorf("usage: outlook-scan get [--save DIR [--attachments]] <message-id>")
	}
	if *withAtts && *save == "" {
		return fmt.Errorf("--attachments requires --save DIR (attachments are binary)")
	}
	c, err := client()
	if err != nil {
		return err
	}
	m, err := c.Get(fs.Arg(0))
	if err != nil {
		return err
	}
	if *asJSON {
		return printJSON(m)
	}
	var atts []graph.Attachment
	if m.HasAttachments {
		if atts, err = c.Attachments(m.ID); err != nil {
			fmt.Fprintln(os.Stderr, "warning: attachment listing failed:", err)
		}
	}
	sv := maybeUnwrapSMIME(c, m, atts)
	if sv != nil {
		atts = applySMIME(m, sv)
	}
	md := render.Message(m, atts)
	if *save != "" {
		path, err := render.Save(*save, m.Received.Local().Format("2006-01-02"), m.Subject, md)
		if err != nil {
			return err
		}
		fmt.Println("saved:", path)
		if *withAtts {
			if sv != nil && !sv.Encrypted {
				return saveInnerParts(sv.Parts, *save)
			}
			return downloadAttachments(c, m.ID, atts, *save)
		}
		return nil
	}
	fmt.Print(md)
	return nil
}

// saveInnerParts writes attachments extracted from an S/MIME container.
func saveInnerParts(parts []smime.Part, dir string) error {
	if len(parts) == 0 {
		fmt.Println("no attachments inside the S/MIME container")
		return nil
	}
	for _, p := range parts {
		path, err := render.SaveRaw(dir, p.Name, p.Data)
		if err != nil {
			return err
		}
		fmt.Println("saved:", path)
	}
	return nil
}

// downloadAttachments writes all file attachments of a message into dir.
func downloadAttachments(c *graph.Client, msgID string, atts []graph.Attachment, dir string) error {
	if len(atts) == 0 {
		fmt.Println("no attachments")
		return nil
	}
	for _, a := range atts {
		if !a.IsFile() {
			fmt.Printf("skipped: %s (not a file attachment: %s)\n", a.Name, a.Type)
			continue
		}
		data, err := c.AttachmentContent(msgID, a.ID)
		if err != nil {
			return fmt.Errorf("download %s: %w", a.Name, err)
		}
		path, err := render.SaveRaw(dir, a.Name, data)
		if err != nil {
			return err
		}
		fmt.Println("saved:", path)
	}
	return nil
}

func cmdAttachments(args []string) error {
	fs := flag.NewFlagSet("attachments", flag.ExitOnError)
	save := fs.String("save", "", "download file attachments to this directory")
	asJSON := fs.Bool("json", false, "raw JSON output")
	fs.Parse(args)
	if fs.NArg() != 1 {
		return fmt.Errorf("usage: outlook-scan attachments [--save DIR] <message-id>")
	}
	c, err := client()
	if err != nil {
		return err
	}
	atts, err := c.Attachments(fs.Arg(0))
	if err != nil {
		return err
	}
	if *asJSON {
		return printJSON(atts)
	}

	// Opaque S/MIME (classic shape: the container is the only attachment):
	// operate on the inner attachments, not the container.
	var sv *smime.Result
	if len(atts) == 1 && smime.IsPKCS7Attachment(atts[0].Name, atts[0].ContentType) {
		if m, err := c.GetMeta(fs.Arg(0)); err == nil {
			m.Body = nil // metadata fetch has no body; force the empty-body check
			sv = maybeUnwrapSMIME(c, m, atts)
		}
	}
	if sv != nil && !sv.Encrypted {
		if *save != "" {
			return saveInnerParts(sv.Parts, *save)
		}
		if len(sv.Parts) == 0 {
			fmt.Println("no attachments inside the S/MIME container")
			return nil
		}
		for i, p := range sv.Parts {
			fmt.Printf("%2d. %s (%s, file, unwrapped from S/MIME)\n",
				i+1, p.Name, render.SizeStr(len(p.Data)))
		}
		return nil
	}
	if sv != nil && sv.Encrypted {
		fmt.Println(smimeEncryptedNote)
		return nil
	}

	if *save != "" {
		return downloadAttachments(c, fs.Arg(0), atts, *save)
	}
	if len(atts) == 0 {
		fmt.Println("no attachments")
		return nil
	}
	for i, a := range atts {
		kind := "file"
		if !a.IsFile() {
			kind = strings.TrimPrefix(a.Type, "#microsoft.graph.")
		}
		fmt.Printf("%2d. %s (%s, %s)\n", i+1, a.Name, render.SizeStr(a.Size), kind)
	}
	return nil
}

func cmdThread(args []string) error {
	fs := flag.NewFlagSet("thread", flag.ExitOnError)
	save := fs.String("save", "", "write markdown to this directory instead of stdout")
	asJSON := fs.Bool("json", false, "raw JSON output")
	fs.Parse(args)
	if fs.NArg() != 1 {
		return fmt.Errorf("usage: outlook-scan thread [--save DIR] <message-id|conversation-id>")
	}
	c, err := client()
	if err != nil {
		return err
	}
	convID := fs.Arg(0)
	// Conversation ids start with AAQk; anything else is treated as a
	// message id and resolved to its conversation.
	if !strings.HasPrefix(convID, "AAQk") {
		m, err := c.GetMeta(convID)
		if err != nil {
			return fmt.Errorf("resolve conversation from message id: %w", err)
		}
		convID = m.ConversationID
	}
	msgs, err := c.Thread(convID)
	if err != nil {
		return err
	}
	if len(msgs) == 0 {
		return fmt.Errorf("no messages found for conversation %s", convID)
	}
	if *asJSON {
		return printJSON(msgs)
	}
	// Unwrap opaque S/MIME members of the thread so they render with content.
	for i := range msgs {
		m := &msgs[i]
		if !m.HasAttachments || (m.Body != nil && strings.TrimSpace(m.Body.Content) != "") {
			continue
		}
		atts, err := c.Attachments(m.ID)
		if err != nil {
			continue
		}
		if sv := maybeUnwrapSMIME(c, m, atts); sv != nil {
			applySMIME(m, sv)
		}
	}
	md := render.Thread(msgs)
	if *save != "" {
		last := msgs[len(msgs)-1]
		path, err := render.Save(*save, last.Received.Local().Format("2006-01-02"),
			"thread-"+last.Subject, md)
		if err != nil {
			return err
		}
		fmt.Println("saved:", path)
		return nil
	}
	fmt.Print(md)
	return nil
}
