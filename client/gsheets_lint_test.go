package client

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"net/url"
	"testing"
)

func TestGSheetsLint_QueryParams(t *testing.T) {
	var gotQuery url.Values

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet {
			t.Fatalf("expected GET, got %s", r.Method)
		}
		if r.URL.Path != "/v0/orgs/org-1/gsheets/sheet-42/lint" {
			t.Fatalf("unexpected path: %s", r.URL.Path)
		}
		gotQuery = r.URL.Query()
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"diagnostics":[],"total":0}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-jwt", "org-1", true)
	c.maxAttempts = 1

	params := url.Values{}
	params.Add("range", "Sheet1!A1:B10")
	params.Add("range", "Sheet2!C1:C20")
	params.Add("skipRule", "D003")
	params.Add("onlyRule", "D004")

	resp, err := c.GSheetsLint("sheet-42", params)
	if err != nil {
		t.Fatalf("GSheetsLint failed: %v", err)
	}
	if resp.Total != 0 {
		t.Fatalf("unexpected total: %d", resp.Total)
	}
	if got := gotQuery["range"]; len(got) != 2 || got[0] != "Sheet1!A1:B10" || got[1] != "Sheet2!C1:C20" {
		t.Fatalf("unexpected range query: %#v", gotQuery["range"])
	}
	if got := gotQuery["skipRule"]; len(got) != 1 || got[0] != "D003" {
		t.Fatalf("unexpected skipRule query: %#v", gotQuery["skipRule"])
	}
	if got := gotQuery["onlyRule"]; len(got) != 1 || got[0] != "D004" {
		t.Fatalf("unexpected onlyRule query: %#v", gotQuery["onlyRule"])
	}
}

func TestGSheetsLint_NotImplementedError(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusNotImplemented)
		fmt.Fprint(w, `{"error":{"code":"NOT_IMPLEMENTED","message":"Rule D032 is not supported for Google Sheets"}}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-jwt", "org-1", true)
	c.maxAttempts = 1

	_, err := c.GSheetsLint("sheet-42", url.Values{"onlyRule": {"D032"}})
	if err == nil {
		t.Fatal("expected error")
	}
	apiErr, ok := err.(*APIError)
	if !ok {
		t.Fatalf("expected *APIError, got %T", err)
	}
	if apiErr.StatusCode != http.StatusNotImplemented {
		t.Fatalf("unexpected status: %d", apiErr.StatusCode)
	}
	if got := err.Error(); got == "" || got == apiErr.Message {
		t.Fatalf("expected friendly error message, got %q", got)
	}
}
