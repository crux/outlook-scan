// outlook-scan is an on-demand mailbox query tool backed by Microsoft
// Graph: list, search, and fetch messages or whole threads as markdown,
// designed for consumption by LLM sessions (stdout-first, --save opt-in).
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/crux/outlook-scan/internal/graph"
	"github.com/crux/outlook-scan/internal/msauth"
	"github.com/crux/outlook-scan/internal/render"
)

func usage() {
	fmt.Fprintln(os.Stderr, `usage: outlook-scan <command> [flags] [args]

commands:
  login                       sign in via device code (one-time; tokens cached)
  status                      show config, session, and signed-in user
  folders                     list mail folders with item counts
  list        [flags]         recent messages (metadata) from a folder
  search      [flags] <query> server-side full-text search across all folders
  get         [flags] <id>    one full message as markdown
  thread      [flags] <id>    whole conversation (by message or conversation id)
  attachments [flags] <id>    list a message's attachments; --save downloads them

flags (list):        --folder NAME  --unread  --since 7d|2w|2026-01-15  --max N  --json
flags (search):      --max N  --json
flags (get):         --save DIR  --attachments (with --save: download files too)  --json
flags (thread):      --save DIR  --json
flags (attachments): --save DIR  --json

Flags come before positional arguments.`)
}

func main() {
	if len(os.Args) < 2 {
		usage()
		os.Exit(2)
	}
	cmds := map[string]func([]string) error{
		"login":   func([]string) error { return cmdLogin() },
		"status":  func([]string) error { return cmdStatus() },
		"folders": func([]string) error { return cmdFolders() },
		"list":        cmdList,
		"search":      cmdSearch,
		"get":         cmdGet,
		"thread":      cmdThread,
		"attachments": cmdAttachments,
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

func cmdLogin() error {
	auth, err := msauth.Load()
	if err != nil {
		return err
	}
	return auth.Login(os.Stdout)
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
	md := render.Message(m, atts)
	if *save != "" {
		path, err := render.Save(*save, m.Received.Local().Format("2006-01-02"), m.Subject, md)
		if err != nil {
			return err
		}
		fmt.Println("saved:", path)
		if *withAtts {
			return downloadAttachments(c, m.ID, atts, *save)
		}
		return nil
	}
	fmt.Print(md)
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
