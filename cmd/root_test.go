package cmd

import (
	"bytes"
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/spf13/pflag"
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

func TestVersionFlag_PrintsHealthVersion(t *testing.T) {
	origVersion := Version
	origRootVersion := rootCmd.Version
	origAPIURL := apiURL
	t.Cleanup(func() {
		Version = origVersion
		rootCmd.Version = origRootVersion
		apiURL = origAPIURL
		resetRootCommandForTest()
	})

	Version = "1.2.3"
	rootCmd.Version = Version
	apiURL = ""

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/health" {
			t.Fatalf("unexpected path: %s", r.URL.Path)
		}
		if got := r.Header.Get("User-Agent"); got != "witan-cli/1.2.3" {
			t.Fatalf("unexpected user-agent header: %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"status":"ok","meta":{"VERSION":"v2.11.1"}}`)
	}))
	defer server.Close()

	var stdout bytes.Buffer
	var stderr bytes.Buffer
	rootCmd.SetOut(&stdout)
	rootCmd.SetErr(&stderr)
	rootCmd.SetArgs([]string{"--api-url", server.URL, "--version"})

	if err := rootCmd.Execute(); err != nil {
		t.Fatalf("Execute returned error: %v", err)
	}

	if got := stdout.String(); got != "witan version 1.2.3\nAPI version: v2.11.1\n" {
		t.Fatalf("unexpected version output: %q", got)
	}
	if got := stderr.String(); got != "" {
		t.Fatalf("expected no stderr output, got %q", got)
	}
}

func TestVersionFlag_PrintsUnavailableWhenHealthFails(t *testing.T) {
	origVersion := Version
	origRootVersion := rootCmd.Version
	origAPIURL := apiURL
	t.Cleanup(func() {
		Version = origVersion
		rootCmd.Version = origRootVersion
		apiURL = origAPIURL
		resetRootCommandForTest()
	})

	Version = "1.2.3"
	rootCmd.Version = Version
	apiURL = ""

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusBadGateway)
	}))
	defer server.Close()

	var stdout bytes.Buffer
	var stderr bytes.Buffer
	rootCmd.SetOut(&stdout)
	rootCmd.SetErr(&stderr)
	rootCmd.SetArgs([]string{"--api-url", server.URL, "--version"})

	if err := rootCmd.Execute(); err != nil {
		t.Fatalf("Execute returned error: %v", err)
	}

	if got := stdout.String(); got != "witan version 1.2.3\nAPI version: unavailable\n" {
		t.Fatalf("unexpected version output: %q", got)
	}
	if got := stderr.String(); got != "" {
		t.Fatalf("expected no stderr output, got %q", got)
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

func TestSkillDocs_UseValidExecExamples(t *testing.T) {
	readSourcePath := filepath.Join("..", "skills", "read-source", "SKILL.md")
	readSource, err := os.ReadFile(readSourcePath)
	if err != nil {
		t.Fatalf("reading %s: %v", readSourcePath, err)
	}
	readSourceText := string(readSource)
	invalidWriteExample := "witan xlsx exec model.xlsx --input-json '...'"
	validWriteExample := `witan xlsx exec model.xlsx --code 'return await xlsx.setCells(wb, [{ address: "Inputs!B3", value: 123 }])'`
	if strings.Contains(readSourceText, invalidWriteExample) {
		t.Fatalf("read-source skill still contains invalid exec example: %s", invalidWriteExample)
	}
	if strings.Contains(readSourceText, "--input-json") {
		t.Fatal("read-source skill should keep the write example self-contained in --code")
	}
	if !strings.Contains(readSourceText, validWriteExample) {
		t.Fatalf("read-source skill should contain valid exec example: %s", validWriteExample)
	}

	xlsxSkillPath := filepath.Join("..", "skills", "xlsx-code-mode", "SKILL.md")
	xlsxSkill, err := os.ReadFile(xlsxSkillPath)
	if err != nil {
		t.Fatalf("reading %s: %v", xlsxSkillPath, err)
	}
	xlsxSkillText := string(xlsxSkill)
	bareCalls := []string{
		"`traceToInputs(wb, outputAddr)`",
		"`setCells` to make the change",
		"* - text: findCells(wb, \"Revenue\")",
		"* - formula search: findCells(wb, \"SUM\", { formulas: true })",
	}
	for _, bad := range bareCalls {
		if strings.Contains(xlsxSkillText, bad) {
			t.Fatalf("xlsx-code-mode skill still contains bare exec API example: %s", bad)
		}
	}
	requiredCalls := []string{
		"`xlsx.traceToInputs(wb, outputAddr)`",
		"`xlsx.setCells` to make the change",
		"* - text: xlsx.findCells(wb, \"Revenue\")",
		"* - formula search: xlsx.findCells(wb, \"SUM\", { formulas: true })",
	}
	for _, good := range requiredCalls {
		if !strings.Contains(xlsxSkillText, good) {
			t.Fatalf("xlsx-code-mode skill should contain documented exec API example: %s", good)
		}
	}
}

func resetRootCommandForTest() {
	rootCmd.SetArgs(nil)
	rootCmd.SetOut(nil)
	rootCmd.SetErr(nil)
	resetCommandFlag(rootCmd.Flags(), "version", "false")
	resetCommandFlag(rootCmd.PersistentFlags(), "api-url", "")
}

func resetCommandFlag(flagSet interface{ Lookup(string) *pflag.Flag }, name, value string) {
	flag := flagSet.Lookup(name)
	if flag == nil {
		return
	}
	_ = flag.Value.Set(value)
	flag.Changed = false
}
