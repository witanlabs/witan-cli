package cmd

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/spf13/cobra"
)

func TestRunPPTXRender_StatelessWritesOutput(t *testing.T) {
	resetPPTXRenderTestGlobals(t)
	filePath, _ := writePresentationForExecTest(t)
	outPath := filepath.Join(t.TempDir(), "slide.png")

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/orgs/org_test/pptx/render" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.URL.Query().Get("slide"); got != "2" {
			t.Fatalf("expected slide=2, got %q", got)
		}
		if got := r.URL.Query().Get("dpr"); got != "3" {
			t.Fatalf("expected dpr=3, got %q", got)
		}
		w.Header().Set("Content-Type", "image/png")
		fmt.Fprint(w, "png bytes")
	}))
	defer server.Close()

	stateless = true
	apiURL = server.URL
	apiKey = "test-key"
	pptxRenderSlide = 2
	pptxRenderDPR = 3
	pptxRenderOutput = outPath

	output, err := captureExecStdout(t, func() error {
		return runPPTXRender(&cobra.Command{}, []string{filePath})
	})
	if err != nil {
		t.Fatalf("runPPTXRender failed: %v", err)
	}
	if !strings.Contains(output, outPath) || !strings.Contains(output, "slide=2 | dpr=3 | image/png") {
		t.Fatalf("unexpected output: %q", output)
	}

	written, err := os.ReadFile(outPath)
	if err != nil {
		t.Fatalf("reading output image: %v", err)
	}
	if string(written) != "png bytes" {
		t.Fatalf("unexpected image bytes: %q", string(written))
	}
}

func TestRunPPTXRender_ValidatesInputs(t *testing.T) {
	resetPPTXRenderTestGlobals(t)

	pptxRenderSlide = 0
	pptxRenderDPR = 1
	err := runPPTXRender(&cobra.Command{}, []string{filepath.Join(t.TempDir(), "deck.pptx")})
	if err == nil || !strings.Contains(err.Error(), "--slide is required") {
		t.Fatalf("unexpected slide validation error: %v", err)
	}

	pptxRenderSlide = 1
	pptxRenderDPR = 4
	err = runPPTXRender(&cobra.Command{}, []string{filepath.Join(t.TempDir(), "deck.pptx")})
	if err == nil || !strings.Contains(err.Error(), "--dpr must be 1-3") {
		t.Fatalf("unexpected dpr validation error: %v", err)
	}

	pptxRenderSlide = 1
	pptxRenderDPR = 1
	err = runPPTXRender(&cobra.Command{}, []string{filepath.Join(t.TempDir(), "deck.pdf")})
	if err == nil || !strings.Contains(err.Error(), "must end in .pptx") {
		t.Fatalf("unexpected extension validation error: %v", err)
	}
}

func resetPPTXRenderTestGlobals(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	origSlide := pptxRenderSlide
	origDPR := pptxRenderDPR
	origOutput := pptxRenderOutput
	origDiff := pptxRenderDiff

	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
		pptxRenderSlide = origSlide
		pptxRenderDPR = origDPR
		pptxRenderOutput = origOutput
		pptxRenderDiff = origDiff
	})

	mockMgmtOrgsServer(t)
	apiKey = ""
	apiURL = ""
	stateless = false
	pptxRenderSlide = 0
	pptxRenderDPR = 1
	pptxRenderOutput = ""
	pptxRenderDiff = ""
}
