// Package graph is a thin Microsoft Graph REST client: auth header,
// retry on 401 (one silent refresh) and on 429/503 (Retry-After).
package graph

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"strconv"
	"time"
)

const base = "https://graph.microsoft.com/v1.0"

type TokenSource interface {
	Token() (string, error)
}

type Client struct {
	Auth TokenSource
	// BodyPreference is sent as the Prefer header's outlook.body-content-type.
	BodyPreference string // "text" (default) or "html"
}

func New(auth TokenSource) *Client {
	return &Client{Auth: auth, BodyPreference: "text"}
}

// GetRaw fetches base+path (or an absolute URL, e.g. an @odata.nextLink)
// and returns the response body bytes.
func (c *Client) GetRaw(path string) ([]byte, error) {
	url := path
	if len(path) > 0 && path[0] == '/' {
		url = base + path
	}
	for attempt := 0; ; attempt++ {
		tok, err := c.Auth.Token()
		if err != nil {
			return nil, err
		}
		req, err := http.NewRequest(http.MethodGet, url, nil)
		if err != nil {
			return nil, err
		}
		req.Header.Set("Authorization", "Bearer "+tok)
		if c.BodyPreference != "" {
			req.Header.Set("Prefer", `outlook.body-content-type="`+c.BodyPreference+`"`)
		}
		resp, err := http.DefaultClient.Do(req)
		if err != nil {
			return nil, err
		}
		body, err := io.ReadAll(resp.Body)
		resp.Body.Close()
		if err != nil {
			return nil, err
		}
		switch {
		case resp.StatusCode == http.StatusOK:
			return body, nil
		case (resp.StatusCode == http.StatusTooManyRequests ||
			resp.StatusCode == http.StatusServiceUnavailable) && attempt < 5:
			wait := 5 * time.Second
			if s := resp.Header.Get("Retry-After"); s != "" {
				if n, err := strconv.Atoi(s); err == nil {
					wait = time.Duration(n) * time.Second
				}
			}
			time.Sleep(wait)
		case resp.StatusCode == http.StatusUnauthorized && attempt == 0:
			continue // Token() refreshes when the cached token has expired
		default:
			return nil, fmt.Errorf("GET %s: HTTP %d: %s", path, resp.StatusCode, truncate(body, 300))
		}
	}
}

// GetJSON fetches like GetRaw and decodes the JSON response into v.
func (c *Client) GetJSON(path string, v any) error {
	body, err := c.GetRaw(path)
	if err != nil {
		return err
	}
	return json.Unmarshal(body, v)
}

// Post sends a JSON body to base+path and decodes any JSON response into out
// (out may be nil). Same auth/retry behavior as GetRaw.
func (c *Client) Post(path string, in, out any) error {
	return c.send(http.MethodPost, path, in, out)
}

// Patch updates a resource at base+path.
func (c *Client) Patch(path string, in, out any) error {
	return c.send(http.MethodPatch, path, in, out)
}

func (c *Client) send(method, path string, in, out any) error {
	raw, err := json.Marshal(in)
	if err != nil {
		return err
	}
	for attempt := 0; ; attempt++ {
		tok, err := c.Auth.Token()
		if err != nil {
			return err
		}
		req, err := http.NewRequest(method, base+path, bytes.NewReader(raw))
		if err != nil {
			return err
		}
		req.Header.Set("Authorization", "Bearer "+tok)
		req.Header.Set("Content-Type", "application/json")
		resp, err := http.DefaultClient.Do(req)
		if err != nil {
			return err
		}
		body, err := io.ReadAll(resp.Body)
		resp.Body.Close()
		if err != nil {
			return err
		}
		switch {
		case resp.StatusCode >= 200 && resp.StatusCode < 300:
			if out != nil && len(body) > 0 {
				return json.Unmarshal(body, out)
			}
			return nil
		case (resp.StatusCode == http.StatusTooManyRequests ||
			resp.StatusCode == http.StatusServiceUnavailable) && attempt < 5:
			wait := 5 * time.Second
			if s := resp.Header.Get("Retry-After"); s != "" {
				if n, err := strconv.Atoi(s); err == nil {
					wait = time.Duration(n) * time.Second
				}
			}
			time.Sleep(wait)
		case resp.StatusCode == http.StatusUnauthorized && attempt == 0:
			continue
		default:
			return fmt.Errorf("%s %s: HTTP %d: %s", method, path, resp.StatusCode, truncate(body, 300))
		}
	}
}

func truncate(b []byte, n int) string {
	if len(b) > n {
		return string(b[:n]) + "..."
	}
	return string(b)
}

// --- typed helpers for the resources the CLI uses ---

type Folder struct {
	ID          string `json:"id"`
	DisplayName string `json:"displayName"`
	TotalItems  int    `json:"totalItemCount"`
	UnreadItems int    `json:"unreadItemCount"`
}

func (c *Client) Folders() ([]Folder, error) {
	var out []Folder
	url := "/me/mailFolders?$select=id,displayName,totalItemCount,unreadItemCount&$top=100"
	for url != "" {
		var page struct {
			Value []Folder `json:"value"`
			Next  string   `json:"@odata.nextLink"`
		}
		if err := c.GetJSON(url, &page); err != nil {
			return nil, err
		}
		out = append(out, page.Value...)
		url = page.Next
	}
	return out, nil
}

type User struct {
	DisplayName       string `json:"displayName"`
	UserPrincipalName string `json:"userPrincipalName"`
}

func (c *Client) Me() (*User, error) {
	u := &User{}
	err := c.GetJSON("/me?$select=displayName,userPrincipalName", u)
	return u, err
}
