package cmd

import (
	"net/http"
	"net/http/httptest"
	"testing"
)

func TestSheetsIntegrationStatus_isActive(t *testing.T) {
	tests := []struct {
		name   string
		status sheetsIntegrationStatus
		want   bool
	}{
		{"not connected", sheetsIntegrationStatus{Connected: false}, false},
		{"active", sheetsIntegrationStatus{Connected: true, Status: "active"}, true},
		{"connected without status", sheetsIntegrationStatus{Connected: true}, false},
		{"needs reauth", sheetsIntegrationStatus{Connected: true, Status: "needs_reauth"}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := tt.status.isActive(); got != tt.want {
				t.Fatalf("isActive() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestSheetsIntegrationStatus_needsReauth(t *testing.T) {
	if !((&sheetsIntegrationStatus{Connected: true, Status: "needs_reauth"}).needsReauth()) {
		t.Fatal("expected needsReauth true")
	}
	if (&sheetsIntegrationStatus{Connected: true, Status: "active"}).needsReauth() {
		t.Fatal("expected needsReauth false for active")
	}
}

func TestGetGoogleSheetsIntegrationStatus(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v0/integrations/google-sheets/status" {
			t.Fatalf("unexpected path: %s", r.URL.Path)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer test-jwt" {
			t.Fatalf("unexpected Authorization: %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write([]byte(`{"connected":true,"expires_at":"2026-06-01T00:00:00Z","status":"active"}`))
	}))
	defer server.Close()

	status, err := getGoogleSheetsIntegrationStatus(server.Client(), server.URL, "test-jwt")
	if err != nil {
		t.Fatalf("getGoogleSheetsIntegrationStatus: %v", err)
	}
	if !status.isActive() {
		t.Fatalf("expected active status, got %+v", status)
	}
	if status.ExpiresAt != "2026-06-01T00:00:00Z" {
		t.Fatalf("expires_at = %q", status.ExpiresAt)
	}
}

func TestGetGoogleSheetsIntegrationStatus_notConnected(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write([]byte(`{"connected":false}`))
	}))
	defer server.Close()

	status, err := getGoogleSheetsIntegrationStatus(server.Client(), server.URL, "test-jwt")
	if err != nil {
		t.Fatalf("getGoogleSheetsIntegrationStatus: %v", err)
	}
	if status.isActive() {
		t.Fatal("expected not active when connected is false")
	}
}

func TestSheetsStatusCheckError_unauthorized(t *testing.T) {
	err := sheetsStatusCheckError(&ManagementAPIError{StatusCode: 401, Code: "unauthorized"})
	if err == nil || err.Error() != "session expired: run 'witan auth login' to re-authenticate" {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestInspectSheetsStatus_fromIntegration(t *testing.T) {
	tests := []struct {
		name        string
		integration sheetsIntegrationStatus
		wantStatus  string
	}{
		{"active", sheetsIntegrationStatus{Connected: true, Status: "active", ExpiresAt: "2026-06-01"}, "connected"},
		{"not connected", sheetsIntegrationStatus{Connected: false}, "not_connected"},
		{"needs reauth", sheetsIntegrationStatus{Connected: true, Status: "needs_reauth"}, "expired"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var report sheetsStatusReport
			switch {
			case tt.integration.needsReauth():
				report = sheetsStatusReport{
					Status:    "expired",
					ExpiresAt: tt.integration.ExpiresAt,
					Error:     "Google authorization has expired or been revoked",
					Hint:      "run 'witan gsheets connect' to reconnect",
				}
			case tt.integration.isActive():
				report = sheetsStatusReport{
					Status:    "connected",
					ExpiresAt: tt.integration.ExpiresAt,
				}
			default:
				report = sheetsStatusReport{Status: "not_connected"}
			}
			if report.Status != tt.wantStatus {
				t.Fatalf("status = %q, want %q", report.Status, tt.wantStatus)
			}
		})
	}
}

func TestValidateSheetsRefRejectsCreateSentinels(t *testing.T) {
	// Real references are valid for existing-sheet ops (lint/render/status).
	for _, ref := range []string{"gs://abc123", "https://docs.google.com/spreadsheets/d/abc123/edit"} {
		if err := validateSheetsRef(ref); err != nil {
			t.Fatalf("%q should be valid: %v", ref, err)
		}
	}
	// The create sentinel must not pass as an existing-sheet ref, in either form:
	// bare "new" fails IsGoogleSheetsURL, while "gs://new" otherwise looks like a
	// valid gs:// ref and needs the explicit sentinel rejection.
	for _, ref := range []string{"new", "gs://new"} {
		if err := validateSheetsRef(ref); err == nil {
			t.Fatalf("validateSheetsRef(%q) = nil, want error", ref)
		}
	}
}

func TestSheetsCreateRejectsPositionalArgs(t *testing.T) {
	if err := sheetsCreateCmd.Args(sheetsCreateCmd, []string{}); err != nil {
		t.Fatalf("create must accept zero args: %v", err)
	}
	if err := sheetsCreateCmd.Args(sheetsCreateCmd, []string{"gs://existing"}); err == nil {
		t.Fatal("create must reject positional args (would silently create a new sheet)")
	}
}

func TestSheetsDisconnectRejectsPositionalArgs(t *testing.T) {
	if err := sheetsDisconnectCmd.Args(sheetsDisconnectCmd, []string{}); err != nil {
		t.Fatalf("disconnect must accept zero args: %v", err)
	}
	if err := sheetsDisconnectCmd.Args(sheetsDisconnectCmd, []string{"gs://sheet"}); err == nil {
		t.Fatal("disconnect must reject positional args (would revoke the whole connection)")
	}
}
