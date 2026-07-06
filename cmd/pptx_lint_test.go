package cmd

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"strings"
	"testing"

	"github.com/spf13/cobra"
)

func TestRunPPTXLint_StatelessPrintsSummary(t *testing.T) {
	resetPPTXLintTestGlobals(t)
	filePath, _ := writePresentationForExecTest(t)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/orgs/org_test/pptx/lint" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.URL.Query()["slide"]; len(got) != 2 || got[0] != "1" || got[1] != "3" {
			t.Fatalf("expected slide=[1 3], got %v", got)
		}
		if got := r.URL.Query().Get("skipRule"); got != "P001" {
			t.Fatalf("expected skipRule=P001, got %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"diagnostics":[],"total":0}`)
	}))
	defer server.Close()

	stateless = true
	apiURL = server.URL
	apiKey = "test-key"
	pptxLintSlides = []int{1, 3}
	pptxLintSkipRule = []string{"P001"}

	output, err := captureExecStdout(t, func() error {
		return runPPTXLint(&cobra.Command{}, []string{filePath})
	})
	if err != nil {
		t.Fatalf("runPPTXLint failed: %v", err)
	}
	if !strings.Contains(output, "0 issues (0 errors, 0 warnings, 0 info)") {
		t.Fatalf("unexpected output: %q", output)
	}
}

func TestRunPPTXLint_ExitsWithCode2OnWarnings(t *testing.T) {
	resetPPTXLintTestGlobals(t)
	filePath, _ := writePresentationForExecTest(t)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"diagnostics":[{"severity":"Warning","ruleId":"P001","message":"Example finding","location":"Slide 2/Title 1 (id 4)","slideNumber":2,"slideId":"256","shapeId":"4","shapeName":"Title 1"}],"total":1}`)
	}))
	defer server.Close()

	stateless = true
	apiURL = server.URL
	apiKey = "test-key"

	output, err := captureExecStdout(t, func() error {
		return runPPTXLint(&cobra.Command{}, []string{filePath})
	})
	exitErr, ok := err.(*ExitError)
	if !ok || exitErr.Code != 2 {
		t.Fatalf("expected ExitError code 2, got %v", err)
	}
	if !strings.Contains(output, "Warning (1):") ||
		!strings.Contains(output, "P001") ||
		!strings.Contains(output, "Slide 2/Title 1 (id 4)") ||
		!strings.Contains(output, "1 issue (0 errors, 1 warning, 0 info)") {
		t.Fatalf("unexpected output: %q", output)
	}
}

func TestRunPPTXLint_ValidatesExtension(t *testing.T) {
	resetPPTXLintTestGlobals(t)

	err := runPPTXLint(&cobra.Command{}, []string{filepath.Join(t.TempDir(), "deck.pdf")})
	if err == nil || !strings.Contains(err.Error(), "must end in .pptx") {
		t.Fatalf("unexpected extension validation error: %v", err)
	}
}

func resetPPTXLintTestGlobals(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	origSlides := pptxLintSlides
	origSkipRule := pptxLintSkipRule
	origOnlyRule := pptxLintOnlyRule

	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
		pptxLintSlides = origSlides
		pptxLintSkipRule = origSkipRule
		pptxLintOnlyRule = origOnlyRule
	})

	mockMgmtOrgsServer(t)
	apiKey = ""
	apiURL = ""
	stateless = false
	pptxLintSlides = nil
	pptxLintSkipRule = nil
	pptxLintOnlyRule = nil
}
