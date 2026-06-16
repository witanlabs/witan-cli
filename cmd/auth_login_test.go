package cmd

import (
	"bufio"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"strings"
	"testing"

	"github.com/witanlabs/witan-cli/config"
)

// captureStdout runs fn with os.Stdout redirected to a pipe and returns what
// was written.
func captureStdout(t *testing.T, fn func()) string {
	t.Helper()
	orig := os.Stdout
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatalf("os.Pipe: %v", err)
	}
	os.Stdout = w
	defer func() { os.Stdout = orig }()

	fn()
	w.Close()
	out, err := io.ReadAll(r)
	if err != nil {
		t.Fatalf("read captured stdout: %v", err)
	}
	return string(out)
}

// TestListOrgs_UsesJWTNotSessionToken verifies that listOrgs is called with a
// JWT obtained from exchangeSessionForJWT, not the raw session token. This
// matches the management API's /v0/orgs endpoint which requires JWT auth.
func TestListOrgs_UsesJWTNotSessionToken(t *testing.T) {
	const (
		sessionToken = "raw-session-token"
		jwtToken     = "jwt-from-exchange"
	)

	mgmt := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch r.URL.Path {
		case "/v0/auth/token":
			// exchangeSessionForJWT: expects session token, returns JWT
			if got := r.Header.Get("Authorization"); got != "Bearer "+sessionToken {
				t.Errorf("/v0/auth/token received unexpected auth header: %q", got)
				w.WriteHeader(http.StatusUnauthorized)
				return
			}
			fmt.Fprintf(w, `{"token":%q}`, jwtToken)

		case "/v0/orgs":
			// listOrgs: must receive the JWT, not the session token
			if got := r.Header.Get("Authorization"); got != "Bearer "+jwtToken {
				t.Errorf("/v0/orgs received unexpected auth header: %q (expected JWT)", got)
				w.WriteHeader(http.StatusUnauthorized)
				return
			}
			fmt.Fprint(w, `{"object":"list","data":[{"id":"org_1","name":"Test Org"}],"has_more":false}`)

		default:
			w.WriteHeader(http.StatusNotFound)
		}
	}))
	defer mgmt.Close()

	// Exchange session token for JWT
	jwt, err := exchangeSessionForJWT(mgmt.URL, sessionToken)
	if err != nil {
		t.Fatalf("exchangeSessionForJWT failed: %v", err)
	}
	if jwt != jwtToken {
		t.Fatalf("expected JWT %q, got %q", jwtToken, jwt)
	}

	// List orgs with the JWT (not the session token)
	orgs, err := listOrgsByJWT(mgmt.URL, jwt)
	if err != nil {
		t.Fatalf("listOrgs failed: %v", err)
	}
	if len(orgs) != 1 || orgs[0].ID != "org_1" {
		t.Fatalf("unexpected orgs: %+v", orgs)
	}
}

func TestSelectOrg_PreferenceMatches(t *testing.T) {
	orgs := []orgEntry{{ID: "org_1", Name: "One"}, {ID: "org_2", Name: "Two"}}
	got, err := selectOrg(orgs, "org_2", "tok", true)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got != "org_2" {
		t.Fatalf("expected org_2, got %q", got)
	}
}

func TestSelectOrg_PreferenceNotFound(t *testing.T) {
	orgs := []orgEntry{{ID: "org_1", Name: "One"}}
	if _, err := selectOrg(orgs, "org_x", "tok", true); err == nil {
		t.Fatal("expected error for unknown org preference")
	}
}

func TestSelectOrg_SingleOrgNoPreference(t *testing.T) {
	orgs := []orgEntry{{ID: "org_only", Name: "Only"}}
	got, err := selectOrg(orgs, "", "tok", true)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got != "org_only" {
		t.Fatalf("expected org_only, got %q", got)
	}
}

// TestSelectOrg_MultiNonInteractiveExits verifies the non-blocking path for an
// agent: with multiple orgs and no preference, selectOrg saves the session
// token (so a re-run with --org can finish) and returns ExitError{Code: 3}
// instead of reading from stdin.
func TestSelectOrg_MultiNonInteractiveExits(t *testing.T) {
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())

	orgs := []orgEntry{{ID: "org_1", Name: "One"}, {ID: "org_2", Name: "Two"}}
	_, err := selectOrg(orgs, "", "saved-token", true)

	var exitErr *ExitError
	if !errors.As(err, &exitErr) {
		t.Fatalf("expected *ExitError, got %v", err)
	}
	if exitErr.Code != 3 {
		t.Fatalf("expected exit code 3, got %d", exitErr.Code)
	}

	cfg, loadErr := config.Load()
	if loadErr != nil {
		t.Fatalf("config.Load failed: %v", loadErr)
	}
	if cfg.SessionToken != "saved-token" {
		t.Fatalf("expected session token saved, got %q", cfg.SessionToken)
	}
	if cfg.SessionOrgID != "" {
		t.Fatalf("expected no org saved, got %q", cfg.SessionOrgID)
	}
}

func TestSelectOrg_NoOrgs(t *testing.T) {
	if _, err := selectOrg(nil, "", "tok", true); err == nil {
		t.Fatal("expected error when no organizations are available")
	}
}

// TestCanResumeOrgSelection guards the fast path that reuses a saved session
// token instead of minting a new device code. It must fire only for an
// incomplete multi-org login (token, no org); a completed session must NOT be
// reused, so a fresh `auth login` always re-authenticates rather than silently
// keeping the previous user's session active.
func TestCanResumeOrgSelection(t *testing.T) {
	pending := config.Config{SessionToken: "tok"}                          // token, no org
	completed := config.Config{SessionToken: "tok", SessionOrgID: "org_1"} // active session

	cases := []struct {
		name           string
		cfg            config.Config
		nonInteractive bool
		orgPref        string
		want           bool
	}{
		{"pending non-interactive with org", pending, true, "org_2", true},
		{"completed session must not resume", completed, true, "org_2", false},
		{"interactive never resumes", pending, false, "org_2", false},
		{"no org preference", pending, true, "", false},
		{"no saved token", config.Config{}, true, "org_2", false},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			if got := canResumeOrgSelection(tc.cfg, tc.nonInteractive, tc.orgPref); got != tc.want {
				t.Fatalf("canResumeOrgSelection = %v, want %v", got, tc.want)
			}
		})
	}
}

// TestJSONOutput_IsParseableJSONL verifies that the two --json emissions a
// single multi-org login can produce (the device-authorization handoff and the
// org_selection_required list) are each one parseable JSON line carrying a type
// discriminator, so a consumer can decode stdout line by line without the two
// objects colliding into one unparseable document.
func TestJSONOutput_IsParseableJSONL(t *testing.T) {
	loginJSON = true
	defer func() { loginJSON = false }()

	dc := &deviceCodeResponse{
		UserCode:                "ABCD1234",
		VerificationURI:         "https://example.test/device",
		VerificationURIComplete: "https://example.test/device?user_code=ABCD1234",
		ExpiresIn:               1800,
	}
	orgs := []orgEntry{{ID: "org_1", Name: "One"}, {ID: "org_2", Name: "Two"}}

	out := captureStdout(t, func() {
		emitHandoff(dc, "ABCD-1234")
		emitOrgChoices(orgs)
	})

	scanner := bufio.NewScanner(strings.NewReader(out))
	var types []string
	for scanner.Scan() {
		line := scanner.Bytes()
		if len(line) == 0 {
			continue
		}
		var obj map[string]any
		if err := json.Unmarshal(line, &obj); err != nil {
			t.Fatalf("line is not parseable JSON: %q: %v", line, err)
		}
		typ, _ := obj["type"].(string)
		if typ == "" {
			t.Fatalf("line missing type discriminator: %q", line)
		}
		types = append(types, typ)
	}
	if err := scanner.Err(); err != nil {
		t.Fatalf("scan: %v", err)
	}

	if len(types) != 2 || types[0] != "device_authorization" || types[1] != "org_selection_required" {
		t.Fatalf("unexpected event types: %v", types)
	}
}

// TestEmitLoginComplete verifies the terminal success event: one parseable
// JSON line carrying type=login_complete plus the resolved org, and nothing at
// all outside --json mode.
func TestEmitLoginComplete(t *testing.T) {
	loginJSON = true
	out := captureStdout(t, func() { emitLoginComplete("a@b.test", "org_9") })
	loginJSON = false

	var obj map[string]any
	if err := json.Unmarshal([]byte(out), &obj); err != nil {
		t.Fatalf("not parseable JSON: %q: %v", out, err)
	}
	if obj["type"] != "login_complete" || obj["org_id"] != "org_9" || obj["email"] != "a@b.test" {
		t.Fatalf("unexpected login_complete payload: %v", obj)
	}

	if silent := captureStdout(t, func() { emitLoginComplete("a@b.test", "org_9") }); silent != "" {
		t.Fatalf("expected no output outside --json, got %q", silent)
	}
}
