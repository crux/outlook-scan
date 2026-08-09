package graph

import (
	"fmt"
	"net/url"
	"sort"
	"strings"
	"time"
)

type Recipient struct {
	EmailAddress struct {
		Name    string `json:"name"`
		Address string `json:"address"`
	} `json:"emailAddress"`
}

func (r Recipient) String() string {
	n, a := r.EmailAddress.Name, r.EmailAddress.Address
	if n != "" && n != a {
		return fmt.Sprintf("%s <%s>", n, a)
	}
	return a
}

type ItemBody struct {
	ContentType string `json:"contentType"`
	Content     string `json:"content"`
}

type Message struct {
	ID             string      `json:"id"`
	Subject        string      `json:"subject"`
	From           *Recipient  `json:"from,omitempty"`
	To             []Recipient `json:"toRecipients,omitempty"`
	Cc             []Recipient `json:"ccRecipients,omitempty"`
	Received       time.Time   `json:"receivedDateTime"`
	IsRead         bool        `json:"isRead"`
	HasAttachments bool        `json:"hasAttachments"`
	ConversationID string      `json:"conversationId"`
	Body           *ItemBody   `json:"body,omitempty"`
}

type Attachment struct {
	ID          string `json:"id"`
	Type        string `json:"@odata.type"` // #microsoft.graph.fileAttachment etc.
	Name        string `json:"name"`
	Size        int    `json:"size"`
	ContentType string `json:"contentType"`
}

// IsFile reports whether the attachment is a plain file (downloadable via
// $value). Item attachments (attached mails) and reference attachments
// (cloud links) are not.
func (a Attachment) IsFile() bool {
	return a.Type == "#microsoft.graph.fileAttachment"
}

const metaSelect = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,conversationId"
const fullSelect = metaSelect + ",body"

// wellKnown maps common folder spellings to Graph well-known folder names.
var wellKnown = map[string]string{
	"inbox": "inbox", "posteingang": "inbox",
	"archive": "archive", "archiv": "archive",
	"sent": "sentitems", "sentitems": "sentitems", "gesendeteelemente": "sentitems",
	"drafts": "drafts", "entwuerfe": "drafts", "entwürfe": "drafts",
	"deleteditems": "deleteditems", "trash": "deleteditems", "gelöschteelemente": "deleteditems",
	"junk": "junkemail", "junkemail": "junkemail", "spam": "junkemail",
	"outbox": "outbox",
}

// resolveFolder turns a user-supplied folder name into a Graph folder ref
// (well-known name or folder id).
func (c *Client) resolveFolder(name string) (string, error) {
	key := strings.ReplaceAll(strings.ToLower(strings.TrimSpace(name)), " ", "")
	if wk, ok := wellKnown[key]; ok {
		return wk, nil
	}
	folders, err := c.Folders()
	if err != nil {
		return "", err
	}
	for _, f := range folders {
		if strings.EqualFold(f.DisplayName, strings.TrimSpace(name)) {
			return f.ID, nil
		}
	}
	names := make([]string, len(folders))
	for i, f := range folders {
		names[i] = f.DisplayName
	}
	return "", fmt.Errorf("folder %q not found; available: %s", name, strings.Join(names, ", "))
}

type ListOpts struct {
	Folder     string // display name or well-known; default Inbox
	UnreadOnly bool
	Since      time.Time // zero = no time filter
	Max        int
}

// List returns message metadata, newest first.
func (c *Client) List(o ListOpts) ([]Message, error) {
	if o.Folder == "" {
		o.Folder = "inbox"
	}
	if o.Max <= 0 {
		o.Max = 25
	}
	ref, err := c.resolveFolder(o.Folder)
	if err != nil {
		return nil, err
	}
	var filters []string
	if o.UnreadOnly {
		filters = append(filters, "isRead eq false")
	}
	if !o.Since.IsZero() {
		filters = append(filters, "receivedDateTime ge "+o.Since.UTC().Format(time.RFC3339))
	}
	q := url.Values{}
	q.Set("$select", metaSelect)
	q.Set("$top", fmt.Sprint(min(o.Max, 100)))
	q.Set("$orderby", "receivedDateTime desc")
	if len(filters) > 0 {
		q.Set("$filter", strings.Join(filters, " and "))
	}
	path := "/me/mailFolders/" + ref + "/messages?" + q.Encode()

	msgs, err := c.pageMessages(path, o.Max)
	if err != nil && strings.Contains(err.Error(), "too complex") {
		// Some filter/orderby combinations exceed Exchange's restriction
		// limits; retry unsorted and sort locally.
		q.Del("$orderby")
		msgs, err = c.pageMessages("/me/mailFolders/"+ref+"/messages?"+q.Encode(), o.Max)
		sort.Slice(msgs, func(i, j int) bool { return msgs[i].Received.After(msgs[j].Received) })
	}
	return msgs, err
}

// Search runs a server-side full-text search over all mail folders.
// Results come back relevance-ranked; Graph caps them at 250.
func (c *Client) Search(query string, maxN int) ([]Message, error) {
	if maxN <= 0 {
		maxN = 25
	}
	q := url.Values{}
	q.Set("$search", `"`+strings.ReplaceAll(query, `"`, `\"`)+`"`)
	q.Set("$select", metaSelect)
	q.Set("$top", fmt.Sprint(min(maxN, 100)))
	return c.pageMessages("/me/messages?"+q.Encode(), maxN)
}

// Get fetches one full message including its body.
func (c *Client) Get(id string) (*Message, error) {
	m := &Message{}
	q := url.Values{}
	q.Set("$select", fullSelect)
	err := c.GetJSON("/me/messages/"+url.PathEscape(id)+"?"+q.Encode(), m)
	return m, err
}

// GetMeta fetches one message's metadata only.
func (c *Client) GetMeta(id string) (*Message, error) {
	m := &Message{}
	q := url.Values{}
	q.Set("$select", metaSelect)
	err := c.GetJSON("/me/messages/"+url.PathEscape(id)+"?"+q.Encode(), m)
	return m, err
}

// Attachments lists attachment metadata for a message.
func (c *Client) Attachments(id string) ([]Attachment, error) {
	var page struct {
		Value []Attachment `json:"value"`
	}
	q := url.Values{}
	q.Set("$select", "id,name,size,contentType")
	err := c.GetJSON("/me/messages/"+url.PathEscape(id)+"/attachments?"+q.Encode(), &page)
	return page.Value, err
}

// AttachmentContent downloads a file attachment's raw bytes.
func (c *Client) AttachmentContent(msgID, attID string) ([]byte, error) {
	return c.GetRaw("/me/messages/" + url.PathEscape(msgID) +
		"/attachments/" + url.PathEscape(attID) + "/$value")
}

// Thread returns all messages of a conversation across all folders,
// oldest first, with full bodies.
func (c *Client) Thread(conversationID string) ([]Message, error) {
	q := url.Values{}
	q.Set("$filter", "conversationId eq '"+conversationID+"'")
	q.Set("$select", fullSelect)
	q.Set("$top", "50")
	msgs, err := c.pageMessages("/me/messages?"+q.Encode(), 500)
	if err != nil {
		return nil, err
	}
	sort.Slice(msgs, func(i, j int) bool { return msgs[i].Received.Before(msgs[j].Received) })
	return msgs, nil
}

func (c *Client) pageMessages(path string, maxN int) ([]Message, error) {
	var out []Message
	for path != "" && len(out) < maxN {
		var page struct {
			Value []Message `json:"value"`
			Next  string    `json:"@odata.nextLink"`
		}
		if err := c.GetJSON(path, &page); err != nil {
			return out, err
		}
		out = append(out, page.Value...)
		path = page.Next
	}
	if len(out) > maxN {
		out = out[:maxN]
	}
	return out, nil
}
