package cmd

import (
	"bytes"
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/config"
)

func TestInspectAuthStatus_NoCredentials(t *testing.T) {
	restoreAuthStatusGlobals(t)

	apiKey = ""
	t.Setenv("WITAN_API_KEY", "")
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())

	report := inspectAuthStatus()

	if report.Status != "unauthenticated" {
		t.Fatalf("expected unauthenticated status, got %q", report.Status)
	}
	if report.ActiveAuth.Type != "none" {
		t.Fatalf("expected no active auth, got %+v", report.ActiveAuth)
	}
	if report.Hint == "" {
		t.Fatal("expected hint when no credentials are configured")
	}
}

func TestInspectAuthStatus_ConfigLoadErrorWithoutCredentials(t *testing.T) {
	restoreAuthStatusGlobals(t)

	apiKey = ""
	t.Setenv("WITAN_API_KEY", "")

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)
	if err := os.Mkdir(filepath.Join(configDir, "config.json"), 0o755); err != nil {
		t.Fatalf("setting up unreadable config path: %v", err)
	}

	report := inspectAuthStatus()

	if report.Status != "unauthenticated" {
		t.Fatalf("expected unauthenticated status, got %+v", report)
	}
	if report.ActiveAuth.Type != "none" {
		t.Fatalf("expected no active auth, got %+v", report.ActiveAuth)
	}
	if report.Error == "" {
		t.Fatal("expected config load error detail")
	}
	if report.Hint != "run `witan auth login` or set `WITAN_API_KEY`" {
		t.Fatalf("unexpected hint: %q", report.Hint)
	}
}

func TestInspectAuthStatus_APIKeyOverridesSavedSession(t *testing.T) {
	restoreAuthStatusGlobals(t)

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)
	if err := config.Save(config.Config{
		SessionToken: "saved-session",
		SessionOrgID: "org_session",
	}); err != nil {
		t.Fatalf("saving config: %v", err)
	}

	mgmt := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/v0/orgs":
			if got := r.Header.Get("Authorization"); got != "ApiKey env-key" {
				t.Fatalf("unexpected auth header: %q", got)
			}
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"object":"list","data":[{"id":"org_api","name":"API Org"}],"has_more":false}`)
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer mgmt.Close()

	t.Setenv("WITAN_MANAGEMENT_API_URL", mgmt.URL)
	t.Setenv("WITAN_API_KEY", "env-key")

	report := inspectAuthStatus()

	if report.Status != "authenticated" {
		t.Fatalf("expected authenticated status, got %+v", report)
	}
	if report.ActiveAuth.Type != "api_key" || report.ActiveAuth.Source != "WITAN_API_KEY" {
		t.Fatalf("unexpected active auth: %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.Validation != "ok" {
		t.Fatalf("expected api key validation ok, got %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.OrgID != "org_api" {
		t.Fatalf("expected org_api, got %+v", report.ActiveAuth)
	}
	if len(report.IgnoredCredentials) != 1 || report.IgnoredCredentials[0].Type != "session" {
		t.Fatalf("expected ignored saved session, got %+v", report.IgnoredCredentials)
	}
}

func TestInspectAuthStatus_APIKeyPrefersCachedOrgID(t *testing.T) {
	restoreAuthStatusGlobals(t)

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)

	cfg := config.Config{}
	cfg.SetOrgIDForAPIKey("env-key", "org_cached")
	if err := config.Save(cfg); err != nil {
		t.Fatalf("saving config: %v", err)
	}

	mgmt := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/v0/orgs":
			if got := r.Header.Get("Authorization"); got != "ApiKey env-key" {
				t.Fatalf("unexpected auth header: %q", got)
			}
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"object":"list","data":[{"id":"org_first","name":"First Org"},{"id":"org_cached","name":"Cached Org"}],"has_more":false}`)
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer mgmt.Close()

	t.Setenv("WITAN_MANAGEMENT_API_URL", mgmt.URL)
	t.Setenv("WITAN_API_KEY", "env-key")

	report := inspectAuthStatus()

	if report.Status != "authenticated" {
		t.Fatalf("expected authenticated status, got %+v", report)
	}
	if report.ActiveAuth.Type != "api_key" || report.ActiveAuth.Source != "WITAN_API_KEY" {
		t.Fatalf("unexpected active auth: %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.Validation != "ok" {
		t.Fatalf("expected api key validation ok, got %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.OrgID != "org_cached" {
		t.Fatalf("expected cached org id, got %+v", report.ActiveAuth)
	}
}

func TestInspectAuthStatus_FlagAPIKeyUsesRawValue(t *testing.T) {
	restoreAuthStatusGlobals(t)

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)

	cfg := config.Config{}
	cfg.SetOrgIDForAPIKey(" key-with-spaces ", "org_raw")
	if err := config.Save(cfg); err != nil {
		t.Fatalf("saving config: %v", err)
	}

	apiKey = " key-with-spaces "
	t.Setenv("WITAN_API_KEY", "env-key")

	mgmt := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/v0/orgs":
			if got := r.Header.Get("Authorization"); got != "ApiKey  key-with-spaces" {
				t.Fatalf("unexpected auth header: %q", got)
			}
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"object":"list","data":[{"id":"org_first","name":"First Org"}],"has_more":false}`)
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer mgmt.Close()

	t.Setenv("WITAN_MANAGEMENT_API_URL", mgmt.URL)

	report := inspectAuthStatus()

	if report.Status != "authenticated" {
		t.Fatalf("expected authenticated status, got %+v", report)
	}
	if report.ActiveAuth.Type != "api_key" || report.ActiveAuth.Source != "--api-key" {
		t.Fatalf("unexpected active auth: %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.Validation != "ok" {
		t.Fatalf("expected api key validation ok, got %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.OrgID != "org_raw" {
		t.Fatalf("expected org_raw, got %+v", report.ActiveAuth)
	}
	if got := report.ActiveAuth.CredentialFingerprint; got != config.HashAPIKey(" key-with-spaces ") {
		t.Fatalf("expected raw-key fingerprint, got %q", got)
	}
	if len(report.IgnoredCredentials) != 1 || report.IgnoredCredentials[0].Source != "WITAN_API_KEY" {
		t.Fatalf("expected env api key to be ignored, got %+v", report.IgnoredCredentials)
	}
}

func TestInspectAuthStatus_InvalidAPIKeyStillMasksSavedSession(t *testing.T) {
	restoreAuthStatusGlobals(t)

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)
	if err := config.Save(config.Config{
		SessionToken: "saved-session",
		SessionOrgID: "org_session",
	}); err != nil {
		t.Fatalf("saving config: %v", err)
	}

	mgmt := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/v0/orgs":
			if got := r.Header.Get("Authorization"); got != "ApiKey bad-key" {
				t.Fatalf("unexpected auth header: %q", got)
			}
			w.WriteHeader(http.StatusUnauthorized)
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer mgmt.Close()

	t.Setenv("WITAN_MANAGEMENT_API_URL", mgmt.URL)
	t.Setenv("WITAN_API_KEY", "bad-key")

	report := inspectAuthStatus()

	if report.Status != "unauthenticated" {
		t.Fatalf("expected unauthenticated status, got %+v", report)
	}
	if report.ActiveAuth.Type != "api_key" || report.ActiveAuth.Validation != "invalid" {
		t.Fatalf("unexpected active auth for invalid api key: %+v", report.ActiveAuth)
	}
	if len(report.IgnoredCredentials) != 1 || report.IgnoredCredentials[0].Type != "session" {
		t.Fatalf("expected ignored saved session, got %+v", report.IgnoredCredentials)
	}
	if !strings.Contains(report.Hint, "saved session") {
		t.Fatalf("expected hint to mention saved session, got %q", report.Hint)
	}
}

func TestInspectAuthStatus_SavedSessionValidatesAndFetchesEmail(t *testing.T) {
	restoreAuthStatusGlobals(t)

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)
	if err := config.Save(config.Config{
		SessionToken: "session-token",
		SessionOrgID: "org_session",
	}); err != nil {
		t.Fatalf("saving config: %v", err)
	}

	mgmt := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch r.URL.Path {
		case "/v0/auth/token":
			if got := r.Header.Get("Authorization"); got != "Bearer session-token" {
				t.Fatalf("unexpected auth header: %q", got)
			}
			fmt.Fprint(w, `{"token":"jwt-token"}`)
		case "/v0/auth/get-session":
			if got := r.Header.Get("Authorization"); got != "Bearer session-token" {
				t.Fatalf("unexpected auth header: %q", got)
			}
			fmt.Fprint(w, `{"user":{"email":"alice@example.com"}}`)
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer mgmt.Close()

	t.Setenv("WITAN_MANAGEMENT_API_URL", mgmt.URL)
	t.Setenv("WITAN_API_KEY", "")

	report := inspectAuthStatus()

	if report.Status != "authenticated" {
		t.Fatalf("expected authenticated status, got %+v", report)
	}
	if report.ActiveAuth.Type != "session" || report.ActiveAuth.Source != savedSessionSource {
		t.Fatalf("unexpected active auth: %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.Validation != "ok" {
		t.Fatalf("expected session validation ok, got %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.UserEmail != "alice@example.com" {
		t.Fatalf("expected alice@example.com, got %+v", report.ActiveAuth)
	}
	if report.ActiveAuth.OrgID != "org_session" {
		t.Fatalf("expected org_session, got %+v", report.ActiveAuth)
	}
}

func TestRunAuthStatus_JSONOutput(t *testing.T) {
	restoreAuthStatusGlobals(t)

	t.Setenv("WITAN_API_KEY", "")
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())

	authStatusJSON = true

	cmd := &cobra.Command{}
	var stdout bytes.Buffer
	cmd.SetOut(&stdout)

	if err := runAuthStatus(cmd, nil); err != nil {
		t.Fatalf("runAuthStatus returned error: %v", err)
	}

	var report authStatusReport
	if err := json.Unmarshal(stdout.Bytes(), &report); err != nil {
		t.Fatalf("parsing json output: %v", err)
	}
	if report.Status != "unauthenticated" || report.ActiveAuth.Type != "none" {
		t.Fatalf("unexpected json report: %+v", report)
	}
}

func restoreAuthStatusGlobals(t *testing.T) {
	t.Helper()

	origAPIKey := apiKey
	origAPIURL := apiURL
	origJSON := authStatusJSON

	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		authStatusJSON = origJSON
	})

	apiKey = ""
	apiURL = ""
	authStatusJSON = false
}
