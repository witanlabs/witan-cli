package cmd

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"testing"
)

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
	orgs, err := listOrgs(mgmt.URL, jwt)
	if err != nil {
		t.Fatalf("listOrgs failed: %v", err)
	}
	if len(orgs) != 1 || orgs[0].ID != "org_1" {
		t.Fatalf("unexpected orgs: %+v", orgs)
	}
}
