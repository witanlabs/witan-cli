package client

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"net/url"
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

func TestPersistentCookieJarDeletesDomainCookieWithNormalizedDomain(t *testing.T) {
	tests := []struct {
		name         string
		setDomain    string
		deleteDomain string
	}{
		{
			name:         "leading dot then no dot",
			setDomain:    ".example.com",
			deleteDomain: "example.com",
		},
		{
			name:         "no dot then leading dot",
			setDomain:    "example.com",
			deleteDomain: ".example.com",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			jar, err := NewPersistentCookieJar(filepath.Join(t.TempDir(), "cookies.json"))
			if err != nil {
				t.Fatalf("NewPersistentCookieJar failed: %v", err)
			}
			u, err := url.Parse("https://api.example.com/v1/exec")
			if err != nil {
				t.Fatalf("url.Parse failed: %v", err)
			}

			jar.SetCookies(u, []*http.Cookie{{
				Name:    "AWSALB",
				Value:   "sticky",
				Domain:  tt.setDomain,
				Path:    "/",
				Expires: time.Now().Add(5 * time.Minute),
			}})
			if got := len(jar.data.Cookies); got != 1 {
				t.Fatalf("expected one persisted cookie, got %d", got)
			}

			jar.SetCookies(u, []*http.Cookie{{
				Name:   "AWSALB",
				Domain: tt.deleteDomain,
				Path:   "/",
				MaxAge: -1,
			}})
			if got := len(jar.data.Cookies); got != 0 {
				t.Fatalf("expected expired cookie to be removed from persisted data, got %d", got)
			}

			reloaded, err := NewPersistentCookieJar(jar.path)
			if err != nil {
				t.Fatalf("reloading persistent cookie jar failed: %v", err)
			}
			if got := reloaded.Cookies(u); len(got) != 0 {
				t.Fatalf("expected expired cookie not to be resurrected, got %v", got)
			}
		})
	}
}
