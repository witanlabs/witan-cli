package client

import (
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"
	"testing"
)

func TestGSheetsExecCreate_RequestShape(t *testing.T) {
	var gotPath string
	var gotBody ExecRequest

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		gotPath = r.URL.Path
		if r.Method != http.MethodPost {
			t.Fatalf("expected POST, got %s", r.Method)
		}
		if err := json.NewDecoder(r.Body).Decode(&gotBody); err != nil {
			t.Fatalf("decoding request body: %v", err)
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"","result":null,"spreadsheet_id":"abc123","url":"https://docs.google.com/spreadsheets/d/abc123"}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-jwt", "org-1", true)
	c.maxAttempts = 1

	resp, err := c.GSheetsExecCreate(ExecRequest{
		Code:  "return 1;",
		Title: "My Sheet",
	})
	if err != nil {
		t.Fatalf("GSheetsExecCreate failed: %v", err)
	}
	if gotPath != "/v0/orgs/org-1/gsheets/new/exec" {
		t.Fatalf("unexpected path: %s", gotPath)
	}
	if gotBody.Code != "return 1;" {
		t.Fatalf("unexpected code: %#v", gotBody.Code)
	}
	if gotBody.Title != "My Sheet" {
		t.Fatalf("unexpected title: %#v", gotBody.Title)
	}
	if !resp.Ok {
		t.Fatalf("expected ok response, got %#v", resp)
	}
	if resp.SpreadsheetID != "abc123" {
		t.Fatalf("unexpected spreadsheet_id: %#v", resp.SpreadsheetID)
	}
	if resp.URL != "https://docs.google.com/spreadsheets/d/abc123" {
		t.Fatalf("unexpected url: %#v", resp.URL)
	}
}

func TestGSheetsExec_ExistingSpreadsheet(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v0/orgs/org-1/gsheets/sheet-42/exec" {
			t.Fatalf("unexpected path: %s", r.URL.Path)
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"","result":true,"spreadsheet_id":"sheet-42","url":"https://docs.google.com/spreadsheets/d/sheet-42"}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-jwt", "org-1", true)
	c.maxAttempts = 1

	resp, err := c.GSheetsExec("sheet-42", ExecRequest{Code: "return true;"})
	if err != nil {
		t.Fatalf("GSheetsExec failed: %v", err)
	}
	if !resp.Ok || resp.SpreadsheetID != "sheet-42" {
		t.Fatalf("unexpected response: %#v", resp)
	}
}
