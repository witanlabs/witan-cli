package cmd

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"

	"github.com/spf13/cobra"
)

func TestRunCalcVerify_StatelessSendsVerifyQueryParam(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	origJSONOutput := jsonOutput
	origCalcRanges := append([]string(nil), calcRanges...)
	origCalcShowTouched := calcShowTouched
	origCalcVerify := calcVerify
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
		jsonOutput = origJSONOutput
		calcRanges = origCalcRanges
		calcShowTouched = origCalcShowTouched
		calcVerify = origCalcVerify
	})

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/xlsx/calc" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.URL.Query().Get("verify"); got != "true" {
			t.Fatalf("expected verify=true, got %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"touched":{},"changed":[],"errors":[]}`)
	}))
	defer server.Close()

	filePath := filepath.Join(t.TempDir(), "book.xlsx")
	if err := os.WriteFile(filePath, []byte("PK\x03\x04test"), 0o644); err != nil {
		t.Fatalf("writing workbook fixture: %v", err)
	}

	apiKey = ""
	apiURL = server.URL
	stateless = true
	jsonOutput = true
	calcRanges = nil
	calcShowTouched = false
	calcVerify = true

	if err := runCalc(&cobra.Command{}, []string{filePath}); err != nil {
		t.Fatalf("runCalc failed: %v", err)
	}
}

func TestRunCalcVerify_FilesBackedSendsVerifyQueryParam(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	origJSONOutput := jsonOutput
	origCalcRanges := append([]string(nil), calcRanges...)
	origCalcShowTouched := calcShowTouched
	origCalcVerify := calcVerify
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
		jsonOutput = origJSONOutput
		calcRanges = origCalcRanges
		calcShowTouched = origCalcShowTouched
		calcVerify = origCalcVerify
	})

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodPost && r.URL.Path == "/v0/files":
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"id":"file_1","object":"file","filename":"book.xlsx","bytes":8,"revision_id":"rev_1","status":"ready"}`)
		case r.Method == http.MethodGet && r.URL.Path == "/v0/files/file_1/xlsx/calc":
			if got := r.URL.Query().Get("verify"); got != "true" {
				t.Fatalf("expected verify=true, got %q", got)
			}
			if got := r.URL.Query().Get("revision"); got != "rev_1" {
				t.Fatalf("expected revision=rev_1, got %q", got)
			}
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"touched":{},"changed":[],"errors":[]}`)
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer server.Close()

	filePath := filepath.Join(t.TempDir(), "book.xlsx")
	if err := os.WriteFile(filePath, []byte("PK\x03\x04test"), 0o644); err != nil {
		t.Fatalf("writing workbook fixture: %v", err)
	}

	apiKey = "test-key"
	apiURL = server.URL
	stateless = false
	jsonOutput = true
	calcRanges = nil
	calcShowTouched = false
	calcVerify = true

	if err := runCalc(&cobra.Command{}, []string{filePath}); err != nil {
		t.Fatalf("runCalc failed: %v", err)
	}
}
