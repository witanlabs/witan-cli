package client

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestNewStatefulClientPersistsAffinityCookies(t *testing.T) {
	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)

	requests := 0
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		requests++
		switch requests {
		case 1:
			if got := r.Header.Get("Cookie"); got != "" {
				t.Fatalf("expected first request without cookies, got %q", got)
			}
			http.SetCookie(w, &http.Cookie{
				Name:    "AWSALB",
				Value:   "sticky",
				Path:    "/",
				Expires: time.Now().Add(5 * time.Minute),
			})
		case 2:
			cookie, err := r.Cookie("AWSALB")
			if err != nil {
				t.Fatalf("expected persisted AWSALB cookie: %v", err)
			}
			if cookie.Value != "sticky" {
				t.Fatalf("unexpected AWSALB cookie value: %q", cookie.Value)
			}
		default:
			t.Fatalf("unexpected request %d", requests)
		}

		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"","result":null}`)
	}))
	defer server.Close()

	first := New(server.URL, "test-key", "", false)
	first.maxAttempts = 1
	if first.HTTPClient.Jar == nil {
		t.Fatal("expected stateful client to use a cookie jar")
	}
	if _, err := first.FilesExec("file_1", "rev_1", ExecRequest{Code: "return 1;"}, false); err != nil {
		t.Fatalf("first FilesExec failed: %v", err)
	}

	cookiePath := filepath.Join(configDir, "cookies.json")
	info, err := os.Stat(cookiePath)
	if err != nil {
		t.Fatalf("expected persisted cookie jar: %v", err)
	}
	if got := info.Mode().Perm(); got != 0o600 {
		t.Fatalf("expected cookies.json mode 0600, got %o", got)
	}

	second := New(server.URL, "test-key", "", false)
	second.maxAttempts = 1
	if _, err := second.FilesExec("file_1", "rev_1", ExecRequest{Code: "return 2;"}, false); err != nil {
		t.Fatalf("second FilesExec failed: %v", err)
	}
}

func TestNewStatelessClientDoesNotUsePersistentCookieJar(t *testing.T) {
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())

	c := New("https://api.test.local", "test-key", "", true)
	if c.HTTPClient.Jar != nil {
		t.Fatal("expected stateless client without cookie jar")
	}
}
