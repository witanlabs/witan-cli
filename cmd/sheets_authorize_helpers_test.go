package cmd

import (
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
)

func TestAuthorizeSheetStart(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v0/integrations/google-sheets/authorize-sheet/start" {
			t.Fatalf("unexpected path: %s", r.URL.Path)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer jwt" {
			t.Fatalf("unexpected Authorization: %q", got)
		}
		body, _ := io.ReadAll(r.Body)
		var req authorizeSheetStartRequest
		if err := json.Unmarshal(body, &req); err != nil {
			t.Fatalf("decoding request: %v", err)
		}
		if req.Spreadsheet != "gs://abc" || req.RedirectURL != "https://cb" {
			t.Fatalf("unexpected request body: %+v", req)
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"picker_url":"https://picker?state=xyz"}`))
	}))
	defer server.Close()

	pickerURL, err := authorizeSheetStart(server.Client(), server.URL, "jwt", "gs://abc", "https://cb")
	if err != nil {
		t.Fatalf("authorizeSheetStart: %v", err)
	}
	if pickerURL != "https://picker?state=xyz" {
		t.Fatalf("picker_url = %q", pickerURL)
	}
}

func TestAuthorizeSheetStartNotConnected(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusNotFound)
		_, _ = w.Write([]byte(`{"error":{"code":"google_sheets_not_connected","message":"not connected"}}`))
	}))
	defer server.Close()

	_, err := authorizeSheetStart(server.Client(), server.URL, "jwt", "gs://abc", "")
	if err == nil || !strings.Contains(err.Error(), "witan gsheets connect") {
		t.Fatalf("expected connect hint, got: %v", err)
	}
}

func TestAuthorizeSheetStatusAuthorized(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v0/integrations/google-sheets/authorize-sheet/status" {
			t.Fatalf("unexpected path: %s", r.URL.Path)
		}
		if got := r.URL.Query().Get("spreadsheet"); got != "gs://abc" {
			t.Fatalf("spreadsheet query = %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"authorized":true}`))
	}))
	defer server.Close()

	authorized, err := authorizeSheetStatus(server.Client(), server.URL, "jwt", "gs://abc")
	if err != nil {
		t.Fatalf("authorizeSheetStatus: %v", err)
	}
	if !authorized {
		t.Fatal("expected authorized=true")
	}
}

func TestAuthorizeSheetStatusTransientStatuses(t *testing.T) {
	for _, code := range []int{
		http.StatusTooManyRequests,
		http.StatusInternalServerError,
		http.StatusBadGateway,
		http.StatusServiceUnavailable,
		http.StatusGatewayTimeout,
	} {
		t.Run(http.StatusText(code), func(t *testing.T) {
			server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
				w.WriteHeader(code)
				_, _ = w.Write([]byte(`{"error":{"code":"x","message":"retry"}}`))
			}))
			defer server.Close()

			_, err := authorizeSheetStatus(server.Client(), server.URL, "jwt", "gs://abc")
			if !errors.Is(err, errSheetsAuthUnavailable) {
				t.Fatalf("status %d: expected errSheetsAuthUnavailable, got: %v", code, err)
			}
		})
	}
}

func TestAuthorizeSheetStatusNotConnectedIsHard(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusNotFound)
		_, _ = w.Write([]byte(`{"error":{"code":"google_sheets_not_connected","message":"x"}}`))
	}))
	defer server.Close()

	_, err := authorizeSheetStatus(server.Client(), server.URL, "jwt", "gs://abc")
	if errors.Is(err, errSheetsAuthUnavailable) {
		t.Fatalf("404 not_connected should be a hard error, got transient: %v", err)
	}
	if err == nil || !strings.Contains(err.Error(), "witan gsheets connect") {
		t.Fatalf("expected connect hint, got: %v", err)
	}
}

func TestIsTransientManagementError(t *testing.T) {
	cases := []struct {
		name string
		err  error
		want bool
	}{
		{"503", &ManagementAPIError{StatusCode: 503}, true},
		{"429", &ManagementAPIError{StatusCode: 429}, true},
		{"401", &ManagementAPIError{StatusCode: 401}, false},
		{"403", &ManagementAPIError{StatusCode: 403}, false},
		{"404", &ManagementAPIError{StatusCode: 404}, false},
		{"network", errors.New("connection refused"), true},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			if got := isTransientManagementError(tc.err); got != tc.want {
				t.Fatalf("isTransientManagementError(%v) = %v, want %v", tc.err, got, tc.want)
			}
		})
	}
}

func TestSheetsAuthorizeError(t *testing.T) {
	tests := []struct {
		code string
		want string
	}{
		{"google_sheets_not_connected", "witan gsheets connect"},
		{"google_sheets_scope_not_granted", "witan gsheets connect"},
		{"google_auth_required", "reconnect"},
		{"forbidden", "witan auth login"},
		{"bad_request", "could not parse"},
	}
	for _, tt := range tests {
		t.Run(tt.code, func(t *testing.T) {
			err := sheetsAuthorizeError(&ManagementAPIError{Code: tt.code, Message: "x"})
			if err == nil || !strings.Contains(err.Error(), tt.want) {
				t.Fatalf("code %s: got %v, want substring %q", tt.code, err, tt.want)
			}
		})
	}
}
