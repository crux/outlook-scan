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

const baseScopes = "openid profile offline_access https://graph.microsoft.com/User.Read "
const readScope = "https://graph.microsoft.com/Mail.Read"
const writeScope = "https://graph.microsoft.com/Mail.ReadWrite"

// scopeString returns the space-separated scope set for the given posture.
// Mail.ReadWrite supersets Mail.Read, so write mode keeps all read access.
func scopeString(write bool) string {
	if write {
		return baseScopes + writeScope
	}
	return baseScopes + readScope
}

// Config identifies the app registration and tenant. Written by setup,
// read from ~/.outlook-scan/config.json.
type Config struct {
	ClientID   string `json:"client_id"`
	TenantID   string `json:"tenant_id"`
	TenantName string `json:"tenant_name,omitempty"`
	// Write, when true, makes this install request Mail.ReadWrite (drafts)
	// instead of read-only Mail.Read. Opt-in per install.
	Write bool `json:"write,omitempty"`
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

// WriteConfig persists the app/tenant identity to the config file.
func WriteConfig(cfg Config) error {
	if err := os.MkdirAll(Dir(), 0o700); err != nil {
		return err
	}
	raw, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(configPath(), raw, 0o600)
}

func Load() (*Auth, error) {
	raw, err := os.ReadFile(configPath())
	if err != nil {
		return nil, fmt.Errorf("no config at %s (run `outlook-scan setup` first): %w", configPath(), err)
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

// scopes returns the scope set for this install's read/write posture.
func (a *Auth) scopes() string { return scopeString(a.cfg.Write) }

// WriteMode reports whether this install is configured for draft writing.
func (a *Auth) WriteMode() bool { return a.cfg.Write }

// HasSession reports whether a refresh token is cached.
func (a *Auth) HasSession() bool { return a.cache.RefreshToken != "" }

// ExpiresAt returns the cached access token's expiry (zero if none).
func (a *Auth) ExpiresAt() time.Time { return a.cache.ExpiresAt }

func endpointFor(tenant, kind string) string {
	return "https://login.microsoftonline.com/" + tenant + "/oauth2/v2.0/" + kind
}

func (a *Auth) endpoint(kind string) string {
	return endpointFor(a.cfg.TenantID, kind)
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
		"scope":         {a.scopes()},
	}
	tok, err := a.tokenRequest(form)
	if err != nil {
		return "", fmt.Errorf("session expired — run `outlook-scan login` (%w)", err)
	}
	return tok, nil
}

// TokenResult is the outcome of a device-code flow.
type TokenResult struct {
	AccessToken  string `json:"access_token"`
	RefreshToken string `json:"refresh_token"`
	Scope        string `json:"scope"`
	ExpiresIn    int    `json:"expires_in"`
}

// RunDeviceCode executes a device-code flow against the given tenant
// ("organizations" or a tenant id) for an arbitrary public client and
// scope string, printing sign-in instructions to w. Nothing is persisted.
func RunDeviceCode(w io.Writer, tenant, clientID, scope string) (*TokenResult, error) {
	form := url.Values{"client_id": {clientID}, "scope": {scope}}
	resp, err := http.PostForm(endpointFor(tenant, "devicecode"), form)
	if err != nil {
		return nil, err
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
		return nil, err
	}
	if dc.Error != "" {
		return nil, fmt.Errorf("device code request: %s: %s", dc.Error, dc.ErrorDesc)
	}
	fmt.Fprintf(w, "To sign in, open %s and enter code: %s\n", dc.VerificationURI, dc.UserCode)
	fmt.Fprintf(w, "(code valid for %d minutes; waiting for sign-in...)\n", dc.ExpiresIn/60)

	interval := time.Duration(max(dc.Interval, 5)) * time.Second
	deadline := time.Now().Add(time.Duration(dc.ExpiresIn) * time.Second)
	form = url.Values{
		"client_id":   {clientID},
		"grant_type":  {"urn:ietf:params:oauth:grant-type:device_code"},
		"device_code": {dc.DeviceCode},
	}
	for time.Now().Before(deadline) {
		time.Sleep(interval)
		resp, err := http.PostForm(endpointFor(tenant, "token"), form)
		if err != nil {
			return nil, err
		}
		body, err := io.ReadAll(resp.Body)
		resp.Body.Close()
		if err != nil {
			return nil, err
		}
		if resp.StatusCode == http.StatusOK {
			tok := &TokenResult{}
			if err := json.Unmarshal(body, tok); err != nil {
				return nil, err
			}
			return tok, nil
		}
		oe := &oauthError{}
		if json.Unmarshal(body, oe) == nil && oe.Code != "" {
			switch oe.Code {
			case "authorization_pending":
				continue
			case "slow_down":
				interval += 5 * time.Second
				continue
			}
			return nil, oe
		}
		return nil, fmt.Errorf("token endpoint: HTTP %d", resp.StatusCode)
	}
	return nil, errors.New("device code expired before sign-in completed")
}

// Login runs the device-code flow using the current read/write posture.
func (a *Auth) Login(w io.Writer) error { return a.LoginWith(w, a.cfg.Write) }

// LoginWith runs the device-code flow requesting the read or write scope set
// and, only on success, persists that posture and caches the tokens. A failed
// or walled sign-in leaves the previous posture and session untouched.
func (a *Auth) LoginWith(w io.Writer, write bool) error {
	tok, err := RunDeviceCode(w, a.cfg.TenantID, a.cfg.ClientID, scopeString(write))
	if err != nil {
		return err
	}
	if write != a.cfg.Write {
		a.cfg.Write = write
		if err := WriteConfig(a.cfg); err != nil {
			return err
		}
	}
	a.cache.AccessToken = tok.AccessToken
	a.cache.RefreshToken = tok.RefreshToken
	a.cache.ExpiresAt = time.Now().Add(time.Duration(tok.ExpiresIn) * time.Second)
	if err := a.save(); err != nil {
		return err
	}
	fmt.Fprintf(w, "Signed in. Tokens cached at %s\n", tokenPath())
	return nil
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
