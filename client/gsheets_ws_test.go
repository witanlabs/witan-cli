package client

import (
	"testing"
)

func TestGSheetsRPCWebSocketURL(t *testing.T) {
	c := New("https://api.example.test", "jwt", "org_1", false)

	rawURL, err := c.GSheetsRPCWebSocketURL()
	if err != nil {
		t.Fatalf("GSheetsRPCWebSocketURL failed: %v", err)
	}
	if rawURL != "wss://api.example.test/v0/orgs/org_1/gsheets/ws" {
		t.Fatalf("unexpected URL: %q", rawURL)
	}
}

func TestGSheetsRPCWebSocketURL_requiresOrg(t *testing.T) {
	c := New("https://api.example.test", "jwt", "", false)

	_, err := c.GSheetsRPCWebSocketURL()
	if err == nil {
		t.Fatal("expected error without org")
	}
}
