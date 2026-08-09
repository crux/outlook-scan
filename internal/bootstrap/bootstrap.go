// Package bootstrap performs the one-time self-service setup: it borrows a
// first-party Microsoft public client for a single ephemeral admin-ish token,
// creates the least-privilege app registration, writes the local config, and
// removes its own bootstrap consent again ("leave no trace").
package bootstrap

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"

	"github.com/crux/outlook-scan/internal/msauth"
)

// DefaultViaClient is "Microsoft Graph Command Line Tools" — a first-party
// public client pre-integrated in every tenant; only its client id is used,
// no local Microsoft tooling is required.
const DefaultViaClient = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

const graphBase = "https://graph.microsoft.com/v1.0"
const graphAppID = "00000003-0000-0000-c000-000000000000" // Microsoft Graph

// Delegated permission ids on the Graph resource (well-known GUIDs).
var scopeIDs = map[string]string{
	"Mail.Read":      "570282fd-fa5c-430d-a7fd-fc8dc98a9dca",
	"User.Read":      "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
	"offline_access": "7427e0e9-2fba-42fe-b0c0-848c9e6a8182",
}

var bootstrapScopes = []string{
	"Application.ReadWrite.All",              // create the app registration
	"DelegatedPermissionGrant.ReadWrite.All", // remove our own consent afterwards
}

type Options struct {
	Name        string // app display name, default "outlook-scan"
	ViaClient   string // bootstrap public client id, default DefaultViaClient
	KeepConsent bool   // skip the self-clean step
}

// Run executes the bootstrap and chains into the normal login flow.
func Run(w io.Writer, o Options) error {
	if o.Name == "" {
		o.Name = "outlook-scan"
	}
	if o.ViaClient == "" {
		o.ViaClient = DefaultViaClient
	}

	scope := ""
	for _, s := range bootstrapScopes {
		scope += "https://graph.microsoft.com/" + s + " "
	}
	tok, err := msauth.RunDeviceCode(w, "organizations", o.ViaClient, strings.TrimSpace(scope))
	if err != nil {
		if strings.Contains(err.Error(), "65001") || strings.Contains(err.Error(), "consent") {
			return fmt.Errorf("your tenant requires admin approval for the setup permissions —"+
				" use the manual portal registration instead (see README): %w", err)
		}
		return err
	}

	claims, err := parseJWTClaims(tok.AccessToken)
	if err != nil {
		return fmt.Errorf("parse token claims: %w", err)
	}
	tenantID, userID := claims["tid"], claims["oid"]
	if tenantID == "" {
		return fmt.Errorf("token contains no tenant id")
	}

	g := &client{token: tok.AccessToken}

	fmt.Fprintf(w, "Looking for app registration %q...\n", o.Name)
	appID, err := findOrCreateApp(g, o.Name)
	if err != nil {
		if strings.Contains(err.Error(), "403") {
			return fmt.Errorf("your account may not register applications in this tenant —"+
				" use the manual portal registration or ask an admin (see README): %w", err)
		}
		return err
	}

	if err := msauth.WriteConfig(msauth.Config{ClientID: appID, TenantID: tenantID}); err != nil {
		return err
	}
	fmt.Fprintf(w, "Config written (app %s, tenant %s).\n", appID, tenantID)

	if !o.KeepConsent {
		fmt.Fprintln(w, "Cleaning up bootstrap permissions...")
		if err := removeBootstrapConsent(g, o.ViaClient, userID); err != nil {
			fmt.Fprintf(w, "warning: could not remove bootstrap consent automatically (%v)\n", err)
			fmt.Fprintf(w, "  You can remove %s manually: Entra portal > Enterprise applications >\n"+
				"  the bootstrap app > Permissions.\n", strings.Join(bootstrapScopes, ", "))
		}
	}

	fmt.Fprintln(w, "Signing in to the new app (first sign-in shows a consent prompt for Mail.Read)...")
	return loginWithRetry(w)
}

// findOrCreateApp returns the appId (client id) of the named registration,
// creating it if absent.
func findOrCreateApp(g *client, name string) (string, error) {
	var page struct {
		Value []struct {
			AppID string `json:"appId"`
		} `json:"value"`
	}
	err := g.do("GET", "/applications?$filter="+
		url.QueryEscape(fmt.Sprintf("displayName eq '%s'", strings.ReplaceAll(name, "'", "''")))+
		"&$select=appId", nil, &page)
	if err != nil {
		return "", err
	}
	if len(page.Value) > 0 {
		return page.Value[0].AppID, nil
	}

	var created struct {
		AppID string `json:"appId"`
	}
	body := map[string]any{
		"displayName":            name,
		"signInAudience":         "AzureADMyOrg",
		"isFallbackPublicClient": true,
		"requiredResourceAccess": []map[string]any{{
			"resourceAppId": graphAppID,
			"resourceAccess": []map[string]string{
				{"id": scopeIDs["Mail.Read"], "type": "Scope"},
				{"id": scopeIDs["User.Read"], "type": "Scope"},
				{"id": scopeIDs["offline_access"], "type": "Scope"},
			},
		}},
	}
	if err := g.do("POST", "/applications", body, &created); err != nil {
		return "", err
	}
	return created.AppID, nil
}

// removeBootstrapConsent strips the bootstrap scopes from the signed-in
// user's own consent grant on the bootstrap client, restoring the record
// to its prior state. Other scopes and other users' grants are untouched.
func removeBootstrapConsent(g *client, viaClient, userID string) error {
	var sps struct {
		Value []struct {
			ID string `json:"id"`
		} `json:"value"`
	}
	err := g.do("GET", "/servicePrincipals?$filter="+
		url.QueryEscape(fmt.Sprintf("appId eq '%s'", viaClient))+"&$select=id", nil, &sps)
	if err != nil {
		return err
	}
	if len(sps.Value) == 0 {
		return nil
	}
	var grants struct {
		Value []struct {
			ID          string `json:"id"`
			ConsentType string `json:"consentType"`
			PrincipalID string `json:"principalId"`
			Scope       string `json:"scope"`
		} `json:"value"`
	}
	err = g.do("GET", "/oauth2PermissionGrants?$filter="+
		url.QueryEscape(fmt.Sprintf("clientId eq '%s'", sps.Value[0].ID)), nil, &grants)
	if err != nil {
		return err
	}
	for _, grant := range grants.Value {
		if grant.ConsentType != "Principal" || grant.PrincipalID != userID {
			continue
		}
		var remaining []string
		removed := false
		for _, s := range strings.Fields(grant.Scope) {
			drop := false
			for _, b := range bootstrapScopes {
				if s == b {
					drop = true
					removed = true
					break
				}
			}
			if !drop {
				remaining = append(remaining, s)
			}
		}
		if removed {
			return g.do("PATCH", "/oauth2PermissionGrants/"+grant.ID,
				map[string]string{"scope": strings.Join(remaining, " ")}, nil)
		}
	}
	return nil
}

// loginWithRetry runs the normal login; fresh app registrations can take a
// moment to propagate, so retry on "app not found" for up to ~1 minute.
func loginWithRetry(w io.Writer) error {
	auth, err := msauth.Load()
	if err != nil {
		return err
	}
	for attempt := 0; ; attempt++ {
		err := auth.Login(w)
		if err == nil {
			return nil
		}
		if attempt < 6 && strings.Contains(err.Error(), "700016") { // app not found yet
			fmt.Fprintln(w, "(new registration still propagating, retrying in 10s...)")
			time.Sleep(10 * time.Second)
			continue
		}
		return err
	}
}

// --- minimal Graph client bound to the ephemeral bootstrap token ---

type client struct {
	token string
}

func (c *client) do(method, path string, in, out any) error {
	var body io.Reader
	if in != nil {
		raw, err := json.Marshal(in)
		if err != nil {
			return err
		}
		body = bytes.NewReader(raw)
	}
	req, err := http.NewRequest(method, graphBase+path, body)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+c.token)
	req.Header.Set("Content-Type", "application/json")
	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	raw, err := io.ReadAll(resp.Body)
	if err != nil {
		return err
	}
	if resp.StatusCode < 200 || resp.StatusCode > 299 {
		if len(raw) > 300 {
			raw = raw[:300]
		}
		return fmt.Errorf("%s %s: HTTP %d: %s", method, path, resp.StatusCode, raw)
	}
	if out != nil && len(raw) > 0 {
		return json.Unmarshal(raw, out)
	}
	return nil
}

// parseJWTClaims decodes the (unverified) payload of a JWT into a string map
// for the handful of claims we need (tid, oid). Verification is unnecessary:
// the token came straight from the token endpoint over TLS.
func parseJWTClaims(jwt string) (map[string]string, error) {
	parts := strings.Split(jwt, ".")
	if len(parts) != 3 {
		return nil, fmt.Errorf("not a JWT")
	}
	payload, err := base64.RawURLEncoding.DecodeString(parts[1])
	if err != nil {
		return nil, err
	}
	var all map[string]any
	if err := json.Unmarshal(payload, &all); err != nil {
		return nil, err
	}
	out := map[string]string{}
	for k, v := range all {
		if s, ok := v.(string); ok {
			out[k] = s
		}
	}
	return out, nil
}
