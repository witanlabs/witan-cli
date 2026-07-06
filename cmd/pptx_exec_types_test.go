package cmd

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"testing"
)

func TestRunPPTXExecTypes_PrintsRawDeclarationsWithoutAuth(t *testing.T) {
	origAPIURL := apiURL
	origAPIKey := apiKey
	origStateless := stateless
	t.Cleanup(func() {
		apiURL = origAPIURL
		apiKey = origAPIKey
		stateless = origStateless
	})

	const wantBody = "// exec types\ndeclare namespace PowerPoint {}\n"

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet || r.URL.Path != "/v0/pptx/exec/types" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.Header.Get("Authorization"); got != "" {
			t.Fatalf("expected no auth header, got %q", got)
		}
		w.Header().Set("Content-Type", "text/plain; charset=utf-8")
		fmt.Fprint(w, wantBody)
	}))
	defer server.Close()

	// No credentials configured on purpose: the command must not require auth.
	apiURL = server.URL
	apiKey = ""
	stateless = false

	output, err := captureExecStdout(t, func() error {
		return runPPTXExecTypes(pptxExecTypesCmd, nil)
	})
	if err != nil {
		t.Fatalf("runPPTXExecTypes: %v", err)
	}
	if output != wantBody {
		t.Fatalf("unexpected output: %q", output)
	}
}

func TestRunPPTXExecTypes_PropagatesAPIError(t *testing.T) {
	origAPIURL := apiURL
	origAPIKey := apiKey
	origStateless := stateless
	t.Cleanup(func() {
		apiURL = origAPIURL
		apiKey = origAPIKey
		stateless = origStateless
	})

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusNotFound)
		fmt.Fprint(w, `{"error":{"code":"not_found","message":"no such route"}}`)
	}))
	defer server.Close()

	apiURL = server.URL
	apiKey = ""
	stateless = false

	_, err := captureExecStdout(t, func() error {
		return runPPTXExecTypes(pptxExecTypesCmd, nil)
	})
	if err == nil {
		t.Fatal("expected error, got nil")
	}
}
