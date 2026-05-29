package cmd

import (
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/spf13/cobra"
)

func TestRunPPTXExec_StatelessSuccessHumanOutputAndNoOverwrite(t *testing.T) {
	resetPPTXExecTestGlobals(t)
	filePath, originalBytes := writePresentationForExecTest(t)
	t.Setenv("WITAN_LOCALE", "")
	t.Setenv("LC_ALL", "")
	t.Setenv("LC_MESSAGES", "")
	t.Setenv("LANG", "en_GB.UTF-8")

	var gotExecCode string
	var gotLocale string
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/orgs/org_test/pptx/exec" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.Header.Get("Accept-Language"); got != "en-GB" {
			t.Fatalf("unexpected Accept-Language header: %q", got)
		}
		if err := r.ParseMultipartForm(10 << 20); err != nil {
			t.Fatalf("parsing multipart form: %v", err)
		}
		var payload map[string]any
		if err := json.Unmarshal([]byte(r.FormValue("exec")), &payload); err != nil {
			t.Fatalf("parsing exec payload: %v", err)
		}
		gotExecCode, _ = payload["code"].(string)
		gotLocale, _ = payload["locale"].(string)

		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"hello\n","result":{"answer":42}}`)
	}))
	defer server.Close()

	stateless = true
	apiURL = server.URL
	apiKey = "test-key"

	cmd := newPPTXExecTestCommand()
	if err := cmd.Flags().Set("code", "return 42;"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}

	output, err := captureExecStdout(t, func() error {
		return runPPTXExec(cmd, []string{filePath})
	})
	if err != nil {
		t.Fatalf("runPPTXExec failed: %v", err)
	}
	if gotExecCode != "return 42;" {
		t.Fatalf("unexpected exec code sent: %q", gotExecCode)
	}
	if gotLocale != "en-GB" {
		t.Fatalf("unexpected locale sent: %q", gotLocale)
	}
	if output != "hello\n{\n  \"answer\": 42\n}\n" {
		t.Fatalf("unexpected output:\n%s", output)
	}

	after, err := os.ReadFile(filePath)
	if err != nil {
		t.Fatalf("reading presentation after exec: %v", err)
	}
	if string(after) != string(originalBytes) {
		t.Fatal("presentation bytes changed, but exec must not overwrite local file")
	}
}

func TestRunPPTXExec_CreateSaveWritesPresentation(t *testing.T) {
	resetPPTXExecTestGlobals(t)
	targetPath := filepath.Join(t.TempDir(), "created.pptx")
	newBytes := []byte{0x50, 0x4b, 0x03, 0x04, 'p', 'p', 't', 'x'}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/orgs/org_test/pptx/exec" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.URL.Query().Get("create"); got != "true" {
			t.Fatalf("expected create=true, got %q", got)
		}
		if got := r.URL.Query().Get("save"); got != "true" {
			t.Fatalf("expected save=true, got %q", got)
		}
		if err := r.ParseMultipartForm(10 << 20); err != nil {
			t.Fatalf("parsing multipart form: %v", err)
		}
		if _, _, err := r.FormFile("file"); err == nil {
			t.Fatal("expected no file part for create mode")
		}
		var payload map[string]any
		if err := json.Unmarshal([]byte(r.FormValue("exec")), &payload); err != nil {
			t.Fatalf("parsing exec payload: %v", err)
		}
		if payload["filename"] != "created.pptx" {
			t.Fatalf("unexpected filename: %#v", payload["filename"])
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprintf(w, `{"ok":true,"stdout":"","result":{"ok":true},"file":"%s"}`, base64.StdEncoding.EncodeToString(newBytes))
	}))
	defer server.Close()

	stateless = false
	apiURL = server.URL
	apiKey = "test-key"

	cmd := newPPTXExecTestCommand()
	if err := cmd.Flags().Set("code", "return true;"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}
	if err := cmd.Flags().Set("create", "true"); err != nil {
		t.Fatalf("setting --create: %v", err)
	}
	if err := cmd.Flags().Set("save", "true"); err != nil {
		t.Fatalf("setting --save: %v", err)
	}

	if _, err := captureExecStdout(t, func() error {
		return runPPTXExec(cmd, []string{targetPath})
	}); err != nil {
		t.Fatalf("runPPTXExec failed: %v", err)
	}

	after, err := os.ReadFile(targetPath)
	if err != nil {
		t.Fatalf("reading created presentation: %v", err)
	}
	if string(after) != string(newBytes) {
		t.Fatalf("presentation bytes were not written: got %v want %v", after, newBytes)
	}
}

func TestRunPPTXExec_StatefulReuploadsOnNotFound(t *testing.T) {
	resetPPTXExecTestGlobals(t)
	filePath, _ := writePresentationForExecTest(t)

	uploadCalls := 0
	execCalls := 0
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodPost && r.URL.Path == "/v0/orgs/org_test/files":
			uploadCalls++
			rev := "rev_1"
			if uploadCalls == 2 {
				rev = "rev_2"
			}
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprintf(w, `{"id":"file_1","object":"file","filename":"deck.pptx","bytes":8,"revision_id":"%s","status":"ready"}`, rev)
		case r.Method == http.MethodPost && r.URL.Path == "/v0/orgs/org_test/files/file_1/pptx/exec":
			execCalls++
			if execCalls == 1 {
				if got := r.URL.Query().Get("revision"); got != "rev_1" {
					t.Fatalf("unexpected first revision: %q", got)
				}
				w.WriteHeader(http.StatusNotFound)
				fmt.Fprint(w, `{"error":{"code":"NOT_FOUND","message":"stale revision"}}`)
				return
			}
			if got := r.URL.Query().Get("revision"); got != "rev_2" {
				t.Fatalf("unexpected retry revision: %q", got)
			}
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"ok":true,"stdout":"done\n","result":{"ok":true}}`)
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer server.Close()

	stateless = false
	apiURL = server.URL
	apiKey = "test-key"

	cmd := newPPTXExecTestCommand()
	if err := cmd.Flags().Set("code", "return true;"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}

	output, err := captureExecStdout(t, func() error {
		return runPPTXExec(cmd, []string{filePath})
	})
	if err != nil {
		t.Fatalf("runPPTXExec failed: %v", err)
	}
	if uploadCalls != 2 || execCalls != 2 {
		t.Fatalf("expected 2 uploads and 2 execs, got %d uploads and %d execs", uploadCalls, execCalls)
	}
	if output != "done\n{\n  \"ok\": true\n}\n" {
		t.Fatalf("unexpected output:\n%s", output)
	}
}

func TestRunPPTXExec_OkFalseReturnsExit1AndSummary(t *testing.T) {
	resetPPTXExecTestGlobals(t)
	filePath, _ := writePresentationForExecTest(t)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":false,"stdout":"","error":{"type":"runtime","code":"EXEC_RUNTIME_ERROR","message":"boom"}}`)
	}))
	defer server.Close()

	stateless = true
	apiURL = server.URL
	apiKey = "test-key"

	cmd := newPPTXExecTestCommand()
	if err := cmd.Flags().Set("code", "throw new Error('boom')"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}

	output, err := captureExecStdout(t, func() error {
		return runPPTXExec(cmd, []string{filePath})
	})
	var exitErr *ExitError
	if err == nil || !errors.As(err, &exitErr) || exitErr.Code != 1 {
		t.Fatalf("expected ExitError code 1, got %v", err)
	}
	if !strings.Contains(output, "runtime (EXEC_RUNTIME_ERROR): boom") {
		t.Fatalf("unexpected output: %q", output)
	}
}

func resetPPTXExecTestGlobals(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	origJSONOutput := pptxJSONOutput
	origExecCode := pptxExecCode
	origExecScript := pptxExecScript
	origExecStdin := pptxExecStdin
	origExecExpr := pptxExecExpr
	origExecInputJSON := pptxExecInputJSON
	origExecLocale := pptxExecLocale
	origExecStdinTimeoutMS := pptxExecStdinTimeoutMS
	origExecTimeoutMS := pptxExecTimeoutMS
	origExecMaxOutputChars := pptxExecMaxOutputChars
	origExecSave := pptxExecSave
	origExecCreate := pptxExecCreate

	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
		pptxJSONOutput = origJSONOutput
		pptxExecCode = origExecCode
		pptxExecScript = origExecScript
		pptxExecStdin = origExecStdin
		pptxExecExpr = origExecExpr
		pptxExecInputJSON = origExecInputJSON
		pptxExecLocale = origExecLocale
		pptxExecStdinTimeoutMS = origExecStdinTimeoutMS
		pptxExecTimeoutMS = origExecTimeoutMS
		pptxExecMaxOutputChars = origExecMaxOutputChars
		pptxExecSave = origExecSave
		pptxExecCreate = origExecCreate
	})

	mockMgmtOrgsServer(t)
	apiKey = ""
	apiURL = ""
	stateless = false
	pptxJSONOutput = false
	pptxExecCode = ""
	pptxExecScript = ""
	pptxExecStdin = false
	pptxExecExpr = ""
	pptxExecInputJSON = ""
	pptxExecLocale = ""
	pptxExecStdinTimeoutMS = defaultExecStdinTimeoutMS
	pptxExecTimeoutMS = 0
	pptxExecMaxOutputChars = 0
	pptxExecSave = false
	pptxExecCreate = false
}

func newPPTXExecTestCommand() *cobra.Command {
	cmd := &cobra.Command{}
	cmd.Flags().StringVar(&pptxExecCode, "code", "", "")
	cmd.Flags().StringVar(&pptxExecScript, "script", "", "")
	cmd.Flags().BoolVar(&pptxExecStdin, "stdin", false, "")
	cmd.Flags().StringVar(&pptxExecExpr, "expr", "", "")
	cmd.Flags().StringVar(&pptxExecInputJSON, "input-json", "", "")
	cmd.Flags().StringVar(&pptxExecLocale, "locale", "", "")
	cmd.Flags().IntVar(&pptxExecStdinTimeoutMS, "stdin-timeout-ms", defaultExecStdinTimeoutMS, "")
	cmd.Flags().IntVar(&pptxExecTimeoutMS, "timeout-ms", 0, "")
	cmd.Flags().IntVar(&pptxExecMaxOutputChars, "max-output-chars", 0, "")
	cmd.Flags().BoolVar(&pptxExecCreate, "create", false, "")
	cmd.Flags().BoolVar(&pptxExecSave, "save", false, "")
	return cmd
}

func writePresentationForExecTest(t *testing.T) (string, []byte) {
	t.Helper()
	path := filepath.Join(t.TempDir(), "deck.pptx")
	content := []byte{0x50, 0x4b, 0x03, 0x04, 'w', 'i', 't', 'a', 'n'}
	if err := os.WriteFile(path, content, 0o644); err != nil {
		t.Fatalf("writing presentation: %v", err)
	}
	return path, content
}
