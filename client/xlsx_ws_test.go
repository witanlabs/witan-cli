package client

import (
	"net/url"
	"testing"
)

func TestFilesXlsxRPCWebSocketURL(t *testing.T) {
	c := New("https://api.example.test", "key", "org_1", false)

	rawURL, err := c.FilesXlsxRPCWebSocketURL("file_1", "rev_1", "Sheet1!A1:B2", "fr-FR")
	if err != nil {
		t.Fatalf("FilesXlsxRPCWebSocketURL failed: %v", err)
	}

	u, err := url.Parse(rawURL)
	if err != nil {
		t.Fatalf("parsing URL: %v", err)
	}
	if u.Scheme != "wss" {
		t.Fatalf("expected wss scheme, got %q", u.Scheme)
	}
	if u.Path != "/v0/orgs/org_1/files/file_1/xlsx/ws" {
		t.Fatalf("unexpected path: %q", u.Path)
	}
	q := u.Query()
	if q.Get("revision") != "rev_1" {
		t.Fatalf("unexpected revision: %q", q.Get("revision"))
	}
	if q.Get("protocol") != "rpc" {
		t.Fatalf("unexpected protocol: %q", q.Get("protocol"))
	}
	if q.Get("hint") != "Sheet1!A1:B2" {
		t.Fatalf("unexpected hint: %q", q.Get("hint"))
	}
	if q.Get("locale") != "fr-FR" {
		t.Fatalf("unexpected locale: %q", q.Get("locale"))
	}
}

func TestStatelessXlsxRPCWebSocketURL(t *testing.T) {
	c := New("http://localhost:3000", "", "", true)

	rawURL, err := c.StatelessXlsxRPCWebSocketURL()
	if err != nil {
		t.Fatalf("StatelessXlsxRPCWebSocketURL failed: %v", err)
	}
	if rawURL != "ws://localhost:3000/v0/xlsx/ws" {
		t.Fatalf("unexpected URL: %q", rawURL)
	}
}
