package cmd

import (
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/coder/websocket"
	"github.com/witanlabs/witan-cli/client"
)

func TestValidateSheetsRPCArgs(t *testing.T) {
	tests := []struct {
		name    string
		create  bool
		args    []string
		wantErr string
	}{
		{name: "existing spreadsheet", args: []string{"gs://abc123"}},
		{name: "missing spreadsheet", args: []string{}, wantErr: "requires exactly 1 spreadsheet reference"},
		{name: "create without arg", create: true},
		{name: "create with new ref", create: true, args: []string{"new"}},
		{name: "create with gs new ref", create: true, args: []string{"gs://new"}},
		{
			name:    "create with real ref",
			create:  true,
			args:    []string{"gs://abc123"},
			wantErr: "--create requires spreadsheet reference 'new' or gs://new",
		},
		{name: "implicit create via new", args: []string{"new"}},
		{
			name:    "too many args on create",
			create:  true,
			args:    []string{"new", "extra"},
			wantErr: "accepts at most 1 spreadsheet reference",
		},
		{name: "invalid ref", args: []string{"not-a-sheet"}, wantErr: "invalid spreadsheet reference"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sheetsRPCCreate = tt.create
			t.Cleanup(func() { sheetsRPCCreate = false })

			err := validateSheetsRPCArgs(sheetsRPCCmd, tt.args)
			if tt.wantErr == "" {
				if err != nil {
					t.Fatalf("unexpected error: %v", err)
				}
				return
			}
			if err == nil {
				t.Fatalf("expected error containing %q", tt.wantErr)
			}
			if !strings.Contains(err.Error(), tt.wantErr) {
				t.Fatalf("error = %q, want substring %q", err.Error(), tt.wantErr)
			}
		})
	}
}

func TestSheetsRPCOpenExistingSendsSpreadsheetIDInInit(t *testing.T) {
	initSeen := make(chan map[string]any, 1)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet || r.URL.Path != "/v0/orgs/org_1/gsheets/ws" {
			t.Errorf("unexpected request: %s %s", r.Method, r.URL.String())
			http.NotFound(w, r)
			return
		}
		serveRPCWebSocket(t, w, r, func(ctx context.Context, conn *websocket.Conn) {
			_, raw, err := conn.Read(ctx)
			if err != nil {
				t.Errorf("reading init: %v", err)
				return
			}
			var init map[string]any
			if err := json.Unmarshal(raw, &init); err != nil {
				t.Errorf("parsing init: %v", err)
				return
			}
			initSeen <- init
			resp := `{"id":"witan-init-1","ok":true,"type":"init","spreadsheet_id":"sheet-42","url":"https://docs.google.com/spreadsheets/d/sheet-42","title":"Budget"}`
			if err := conn.Write(ctx, websocket.MessageText, []byte(resp)); err != nil {
				t.Errorf("writing init response: %v", err)
			}
		})
	}))
	defer server.Close()

	c := client.New(server.URL, "jwt", "org_1", false)
	session, err := openSheetsRPCSession(context.Background(), c, sheetsRPCConnectParams{
		SpreadsheetID: "sheet-42",
		Locale:        "en-US",
	})
	if err != nil {
		t.Fatalf("openSheetsRPCSession failed: %v", err)
	}
	session.close()

	init := <-initSeen
	if init["type"] != "init" || init["id"] != "witan-init-1" {
		t.Fatalf("unexpected init envelope: %#v", init)
	}
	if init["spreadsheet_id"] != "sheet-42" {
		t.Fatalf("expected spreadsheet_id in init, got %#v", init)
	}
	if init["locale"] != "en-US" {
		t.Fatalf("unexpected locale: %#v", init)
	}
	if _, ok := init["create"]; ok {
		t.Fatalf("open init must not include create: %#v", init)
	}
	if session.spreadsheetID != "sheet-42" {
		t.Fatalf("session spreadsheet_id = %q", session.spreadsheetID)
	}
	if session.url == "" {
		t.Fatal("expected session url from init response")
	}
}

func TestSheetsRPCCreateSendsCreateInit(t *testing.T) {
	initSeen := make(chan map[string]any, 1)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet || r.URL.Path != "/v0/orgs/org_1/gsheets/ws" {
			t.Errorf("unexpected request: %s %s", r.Method, r.URL.String())
			http.NotFound(w, r)
			return
		}
		serveRPCWebSocket(t, w, r, func(ctx context.Context, conn *websocket.Conn) {
			_, raw, err := conn.Read(ctx)
			if err != nil {
				t.Errorf("reading init: %v", err)
				return
			}
			var init map[string]any
			if err := json.Unmarshal(raw, &init); err != nil {
				t.Errorf("parsing init: %v", err)
				return
			}
			initSeen <- init
			resp := `{"id":"witan-init-1","ok":true,"type":"init","spreadsheet_id":"new-sheet","url":"https://docs.google.com/spreadsheets/d/new-sheet","title":"Q4 Model"}`
			if err := conn.Write(ctx, websocket.MessageText, []byte(resp)); err != nil {
				t.Errorf("writing init response: %v", err)
			}
		})
	}))
	defer server.Close()

	c := client.New(server.URL, "jwt", "org_1", false)
	session, err := openSheetsRPCSession(context.Background(), c, sheetsRPCConnectParams{
		Create: true,
		Title:  "Q4 Model",
		Locale: "fr-FR",
	})
	if err != nil {
		t.Fatalf("openSheetsRPCSession failed: %v", err)
	}
	session.close()

	init := <-initSeen
	if init["create"] != true {
		t.Fatalf("expected create init, got %#v", init)
	}
	if init["title"] != "Q4 Model" {
		t.Fatalf("unexpected title: %#v", init)
	}
	if init["locale"] != "fr-FR" {
		t.Fatalf("unexpected locale: %#v", init)
	}
	if _, ok := init["spreadsheet_id"]; ok {
		t.Fatalf("create init must not include spreadsheet_id: %#v", init)
	}
	if session.spreadsheetID != "new-sheet" {
		t.Fatalf("session spreadsheet_id = %q", session.spreadsheetID)
	}
	if session.title != "Q4 Model" {
		t.Fatalf("session title = %q", session.title)
	}
}

func TestFormatSheetsRPCInitError(t *testing.T) {
	tests := []struct {
		name    string
		resp    sheetsRPCInitResponse
		wantErr string
	}{
		{
			name: "google auth",
			resp: sheetsRPCInitResponse{Code: "google_auth_required", Message: "nope"},
			wantErr: "Google Sheets requires authorization",
		},
		{
			name: "not found",
			resp: sheetsRPCInitResponse{Code: "google_sheets_not_found", Message: "missing"},
			wantErr: "spreadsheet not found",
		},
		{
			name: "invalid init",
			resp: sheetsRPCInitResponse{Code: "INVALID_INIT", Message: "bad fields"},
			wantErr: "INVALID_INIT: bad fields",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := formatSheetsRPCInitError(&tt.resp)
			if err == nil {
				t.Fatal("expected error")
			}
			if !strings.Contains(err.Error(), tt.wantErr) {
				t.Fatalf("error = %q, want substring %q", err.Error(), tt.wantErr)
			}
		})
	}
}
