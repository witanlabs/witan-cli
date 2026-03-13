package cmd

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestResolveStateless_ForcesWithoutCredentials(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
	})

	apiKey = ""
	apiURL = ""
	stateless = false

	t.Setenv("WITAN_API_KEY", "")
	t.Setenv("WITAN_STATELESS", "")
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())

	if !resolveStateless() {
		t.Fatal("expected stateless mode to be forced when no credentials are available")
	}

	key, _, err := resolveAuth()
	if err != nil {
		t.Fatalf("resolveAuth returned error: %v", err)
	}
	if key != "" {
		t.Fatalf("expected empty API key in forced stateless mode, got %q", key)
	}
}

func TestResolveStateless_DoesNotForceWithAPIKey(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
	})

	apiKey = ""
	apiURL = ""
	stateless = false

	// Stand up a mock mgmt API that returns a single org for /v0/orgs
	mgmtServer := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"object":"list","data":[{"id":"org_1","name":"Test Org"}],"has_more":false}`)
	}))
	defer mgmtServer.Close()
	t.Setenv("WITAN_MANAGEMENT_API_URL", mgmtServer.URL)

	t.Setenv("WITAN_API_KEY", "test-key")
	t.Setenv("WITAN_STATELESS", "")
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())

	if resolveStateless() {
		t.Fatal("expected stateful mode when API key is present and stateless is not requested")
	}

	key, orgID, err := resolveAuth()
	if err != nil {
		t.Fatalf("resolveAuth returned error: %v", err)
	}
	if key != "test-key" {
		t.Fatalf("expected API key from environment, got %q", key)
	}
	if orgID != "org_1" {
		t.Fatalf("expected org_1, got %q", orgID)
	}
}

func TestResolveStateless_ForcesWhenConfigLoadErrors(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
	})

	apiKey = ""
	apiURL = ""
	stateless = false

	t.Setenv("WITAN_API_KEY", "")
	t.Setenv("WITAN_STATELESS", "")

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)
	if err := os.Mkdir(filepath.Join(configDir, "config.json"), 0o755); err != nil {
		t.Fatalf("setup invalid config path: %v", err)
	}

	if !resolveStateless() {
		t.Fatal("expected stateless mode to be forced when config cannot be loaded")
	}
}

func TestResolveAPIKey_AllowsStatelessFallbackWhenConfigLoadErrors(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
	})

	apiKey = ""
	apiURL = ""
	stateless = false

	t.Setenv("WITAN_API_KEY", "")
	t.Setenv("WITAN_STATELESS", "")

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)
	if err := os.Mkdir(filepath.Join(configDir, "config.json"), 0o755); err != nil {
		t.Fatalf("setup invalid config path: %v", err)
	}

	key, _, err := resolveAuth()
	if err != nil {
		t.Fatalf("expected stateless fallback, got error: %v", err)
	}
	if key != "" {
		t.Fatalf("expected empty API key in forced stateless mode, got %q", key)
	}
}

func TestNewAPIClient_SetsVersionedUserAgent(t *testing.T) {
	origVersion := Version
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		Version = origVersion
		apiURL = origAPIURL
		stateless = origStateless
	})

	Version = "1.2.3"
	apiURL = "https://api.witanlabs.test"
	stateless = true

	c := newAPIClient("test-key", "")
	if got := c.UserAgent; got != "witan-cli/1.2.3" {
		t.Fatalf("unexpected client user-agent: %q", got)
	}
}

func TestExchangeSessionForJWT_SendsUserAgentHeader(t *testing.T) {
	origVersion := Version
	t.Cleanup(func() {
		Version = origVersion
	})
	Version = "9.9.9"

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.Header.Get("User-Agent"); got != "witan-cli/9.9.9" {
			t.Fatalf("unexpected user-agent header: %q", got)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer test-session" {
			t.Fatalf("unexpected auth header: %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"token":"jwt-token"}`)
	}))
	defer server.Close()

	token, err := exchangeSessionForJWT(server.URL, "test-session")
	if err != nil {
		t.Fatalf("exchangeSessionForJWT returned error: %v", err)
	}
	if token != "jwt-token" {
		t.Fatalf("unexpected token: %q", token)
	}
}

// mockMgmtOrgsServer starts a mock management API that returns a single org
// for GET /v0/orgs and sets WITAN_MANAGEMENT_API_URL. Call t.Cleanup to tear it down.
func mockMgmtOrgsServer(t *testing.T) {
	t.Helper()
	mgmt := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path == "/v0/orgs" {
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"object":"list","data":[{"id":"org_test","name":"Test Org"}],"has_more":false}`)
			return
		}
		w.WriteHeader(http.StatusNotFound)
	}))
	t.Cleanup(mgmt.Close)
	t.Setenv("WITAN_MANAGEMENT_API_URL", mgmt.URL)
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())
}

func TestHelpExamples_UseDocumentedExecAPI(t *testing.T) {
	badExample := `wb.sheet("Summary").cell("A1").value`
	goodExample := `await xlsx.readCell(wb, "Summary!A1")`

	for _, tc := range []struct {
		name string
		text string
	}{
		{name: "root", text: rootCmd.Long},
		{name: "xlsx", text: xlsxCmd.Long},
		{name: "xlsx exec", text: xlsxExecCmd.Long},
	} {
		if strings.Contains(tc.text, badExample) {
			t.Fatalf("%s help still contains undocumented exec API example: %s", tc.name, badExample)
		}
		if !strings.Contains(tc.text, goodExample) {
			t.Fatalf("%s help should contain documented exec API example: %s", tc.name, goodExample)
		}
	}
}
