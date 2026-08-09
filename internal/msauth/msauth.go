// Package msauth implements the OAuth2 device-code flow against Microsoft
// Entra ID for a public client, with a local token cache and silent refresh.
package msauth

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"
)

const scopes = "openid profile offline_access " +
	"https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/User.Read"

// Config identifies the app registration and tenant. Written by setup,
// read from ~/.outlook-scan/config.json.
type Config struct {
	ClientID   string `json:"client_id"`
	TenantID   string `json:"tenant_id"`
	TenantName string `json:"tenant_name,omitempty"`
}

type tokenCache struct {
	AccessToken  string    `json:"access_token"`
	RefreshToken string    `json:"refresh_token"`
	ExpiresAt    time.Time `json:"expires_at"`
}

type Auth struct {
	cfg   Config
	cache tokenCache
}

func Dir() string {
	home, err := os.UserHomeDir()
	if err != nil {
		return ".outlook-scan"
	}
	return filepath.Join(home, ".outlook-scan")
}

func configPath() string { return filepath.Join(Dir(), "config.json") }
func tokenPath() string  { return filepath.Join(Dir(), "token.json") }

func Load() (*Auth, error) {
	raw, err := os.ReadFile(configPath())
	if err != nil {
		return nil, fmt.Errorf("no config at %s (run setup first): %w", configPath(), err)
	}
	a := &Auth{}
	if err := json.Unmarshal(raw, &a.cfg); err != nil {
		return nil, fmt.Errorf("parse %s: %w", configPath(), err)
	}
	if a.cfg.ClientID == "" || a.cfg.TenantID == "" {
		return nil, fmt.Errorf("%s missing client_id or tenant_id", configPath())
	}
	if raw, err := os.ReadFile(tokenPath()); err == nil {
		_ = json.Unmarshal(raw, &a.cache) // corrupt cache just means re-login
	}
	return a, nil
}

func (a *Auth) Config() Config { return a.cfg }

// HasSession reports whether a refresh token is cached.
func (a *Auth) HasSession() bool { return a.cache.RefreshToken != "" }

// ExpiresAt returns the cached access token's expiry (zero if none).
func (a *Auth) ExpiresAt() time.Time { return a.cache.ExpiresAt }

func (a *Auth) endpoint(kind string) string {
	return "https://login.microsoftonline.com/" + a.cfg.TenantID + "/oauth2/v2.0/" + kind
}

// Token returns a valid access token, refreshing silently when needed.
func (a *Auth) Token() (string, error) {
	if a.cache.AccessToken != "" && time.Now().Add(2*time.Minute).Before(a.cache.ExpiresAt) {
		return a.cache.AccessToken, nil
	}
	if a.cache.RefreshToken == "" {
		return "", errors.New("no session — run `outlook-scan login`")
	}
	form := url.Values{
		"client_id":     {a.cfg.ClientID},
		"grant_type":    {"refresh_token"},
		"refresh_token": {a.cache.RefreshToken},
		"scope":         {scopes},
	}
	tok, err := a.tokenRequest(form)
	if err != nil {
		return "", fmt.Errorf("session expired — run `outlook-scan login` (%w)", err)
	}
	return tok, nil
}

// Login runs the device-code flow, printing instructions to w.
func (a *Auth) Login(w io.Writer) error {
	form := url.Values{"client_id": {a.cfg.ClientID}, "scope": {scopes}}
	resp, err := http.PostForm(a.endpoint("devicecode"), form)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	var dc struct {
		DeviceCode      string `json:"device_code"`
		UserCode        string `json:"user_code"`
		VerificationURI string `json:"verification_uri"`
		ExpiresIn       int    `json:"expires_in"`
		Interval        int    `json:"interval"`
		Error           string `json:"error"`
		ErrorDesc       string `json:"error_description"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&dc); err != nil {
		return err
	}
	if dc.Error != "" {
		return fmt.Errorf("device code request: %s: %s", dc.Error, dc.ErrorDesc)
	}
	fmt.Fprintf(w, "To sign in, open %s and enter code: %s\n", dc.VerificationURI, dc.UserCode)
	fmt.Fprintf(w, "(code valid for %d minutes; waiting for sign-in...)\n", dc.ExpiresIn/60)

	interval := time.Duration(max(dc.Interval, 5)) * time.Second
	deadline := time.Now().Add(time.Duration(dc.ExpiresIn) * time.Second)
	form = url.Values{
		"client_id":   {a.cfg.ClientID},
		"grant_type":  {"urn:ietf:params:oauth:grant-type:device_code"},
		"device_code": {dc.DeviceCode},
	}
	for time.Now().Before(deadline) {
		time.Sleep(interval)
		_, err := a.tokenRequest(form)
		if err == nil {
			fmt.Fprintf(w, "Signed in. Tokens cached at %s\n", tokenPath())
			return nil
		}
		var oe *oauthError
		if errors.As(err, &oe) {
			switch oe.Code {
			case "authorization_pending":
				continue
			case "slow_down":
				interval += 5 * time.Second
				continue
			}
		}
		return err
	}
	return errors.New("device code expired before sign-in completed")
}

type oauthError struct {
	Code string `json:"error"`
	Desc string `json:"error_description"`
}

func (e *oauthError) Error() string {
	d := e.Desc
	if i := strings.IndexByte(d, '\n'); i > 0 {
		d = d[:i]
	}
	return e.Code + ": " + d
}

// tokenRequest posts to the token endpoint, updates and persists the cache.
func (a *Auth) tokenRequest(form url.Values) (string, error) {
	resp, err := http.PostForm(a.endpoint("token"), form)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return "", err
	}
	if resp.StatusCode != http.StatusOK {
		oe := &oauthError{}
		if json.Unmarshal(body, oe) == nil && oe.Code != "" {
			return "", oe
		}
		return "", fmt.Errorf("token endpoint: HTTP %d", resp.StatusCode)
	}
	var tok struct {
		AccessToken  string `json:"access_token"`
		RefreshToken string `json:"refresh_token"`
		ExpiresIn    int    `json:"expires_in"`
	}
	if err := json.Unmarshal(body, &tok); err != nil {
		return "", err
	}
	a.cache.AccessToken = tok.AccessToken
	a.cache.ExpiresAt = time.Now().Add(time.Duration(tok.ExpiresIn) * time.Second)
	if tok.RefreshToken != "" { // rotation: keep old one if none returned
		a.cache.RefreshToken = tok.RefreshToken
	}
	return tok.AccessToken, a.save()
}

func (a *Auth) save() error {
	if err := os.MkdirAll(Dir(), 0o700); err != nil {
		return err
	}
	raw, err := json.Marshal(a.cache)
	if err != nil {
		return err
	}
	return os.WriteFile(tokenPath(), raw, 0o600)
}
