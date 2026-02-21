package cmd

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"testing"

	"github.com/spf13/cobra"
)

func TestResolveExecCodeSource_Exclusivity(t *testing.T) {
	resetExecTestGlobals(t)

	t.Run("none selected returns error", func(t *testing.T) {
		cmd := newExecTestCommand()
		_, err := resolveExecCodeSource(cmd, strings.NewReader(""))
		if err == nil || !strings.Contains(err.Error(), "exactly one of --code, --script, --stdin, or --expr is required") {
			t.Fatalf("unexpected error: %v", err)
		}
	})

	t.Run("multiple selected returns error", func(t *testing.T) {
		cmd := newExecTestCommand()
		if err := cmd.Flags().Set("code", "return 1"); err != nil {
			t.Fatalf("setting --code: %v", err)
		}
		if err := cmd.Flags().Set("expr", "1+1"); err != nil {
			t.Fatalf("setting --expr: %v", err)
		}

		_, err := resolveExecCodeSource(cmd, strings.NewReader(""))
		if err == nil || !strings.Contains(err.Error(), "mutually exclusive") {
			t.Fatalf("unexpected error: %v", err)
		}
	})

	t.Run("single source selected passes", func(t *testing.T) {
		cmd := newExecTestCommand()
		if err := cmd.Flags().Set("code", "return 7;"); err != nil {
			t.Fatalf("setting --code: %v", err)
		}

		code, err := resolveExecCodeSource(cmd, strings.NewReader(""))
		if err != nil {
			t.Fatalf("resolveExecCodeSource failed: %v", err)
		}
		if code != "return 7;" {
			t.Fatalf("unexpected code: %q", code)
		}
	})
}

func TestResolveExecCodeSource_ExprWrapsExactly(t *testing.T) {
	resetExecTestGlobals(t)
	cmd := newExecTestCommand()
	if err := cmd.Flags().Set("expr", `input.value + 1`); err != nil {
		t.Fatalf("setting --expr: %v", err)
	}

	code, err := resolveExecCodeSource(cmd, strings.NewReader(""))
	if err != nil {
		t.Fatalf("resolveExecCodeSource failed: %v", err)
	}
	if code != "return (input.value + 1);" {
		t.Fatalf("unexpected wrapped expression: %q", code)
	}
}

func TestResolveExecCodeSource_ScriptAndStdin(t *testing.T) {
	resetExecTestGlobals(t)

	t.Run("script reads file content", func(t *testing.T) {
		cmd := newExecTestCommand()
		scriptPath := filepath.Join(t.TempDir(), "script.js")
		if err := os.WriteFile(scriptPath, []byte("console.log('x')"), 0o644); err != nil {
			t.Fatalf("writing script: %v", err)
		}
		if err := cmd.Flags().Set("script", scriptPath); err != nil {
			t.Fatalf("setting --script: %v", err)
		}

		code, err := resolveExecCodeSource(cmd, strings.NewReader(""))
		if err != nil {
			t.Fatalf("resolveExecCodeSource failed: %v", err)
		}
		if code != "console.log('x')" {
			t.Fatalf("unexpected script content: %q", code)
		}
	})

	t.Run("stdin reads code bytes only", func(t *testing.T) {
		cmd := newExecTestCommand()
		if err := cmd.Flags().Set("stdin", "true"); err != nil {
			t.Fatalf("setting --stdin: %v", err)
		}

		code, err := resolveExecCodeSource(cmd, strings.NewReader("return input;\n"))
		if err != nil {
			t.Fatalf("resolveExecCodeSource failed: %v", err)
		}
		if code != "return input;\n" {
			t.Fatalf("unexpected stdin code: %q", code)
		}
	})
}

func TestParseExecInput(t *testing.T) {
	resetExecTestGlobals(t)

	input, err := parseExecInput("", false)
	if err != nil {
		t.Fatalf("parseExecInput default failed: %v", err)
	}
	inputObj, ok := input.(map[string]any)
	if !ok || len(inputObj) != 0 {
		t.Fatalf("expected default empty object, got %#v", input)
	}

	input, err = parseExecInput(`{"threshold":10}`, true)
	if err != nil {
		t.Fatalf("parseExecInput JSON failed: %v", err)
	}
	obj, ok := input.(map[string]any)
	if !ok || obj["threshold"] != float64(10) {
		t.Fatalf("unexpected parsed input: %#v", input)
	}

	_, err = parseExecInput(`{"threshold":`, true)
	if err == nil || !strings.Contains(err.Error(), "invalid --input-json") {
		t.Fatalf("expected JSON parse error, got: %v", err)
	}
}

func TestXlsxExecHelp_ContractSectionsPresent(t *testing.T) {
	required := []string{
		"Contract:",
		"Inputs:",
		"Defaults:",
		"Output:",
		"Exit codes:",
		"--json prints the full response envelope.",
		`{"ok":true,"stdout":"...","result":<json>`,
		`{"ok":false,"stdout":"...","error":{"type":"...","code":"...","message":"..."}}`,
		"--input-json is omitted, input defaults to {}.",
		"--timeout-ms=0 means no explicit timeout override.",
		"--max-output-chars=0 means no explicit stdout cap override.",
	}

	for _, needle := range required {
		if !strings.Contains(xlsxExecCmd.Long, needle) {
			t.Fatalf("expected help text to contain %q", needle)
		}
	}

	disallowed := []string{
		"/v0/xlsx/exec",
		"/v0/files/:id/xlsx/exec",
	}
	if slices.ContainsFunc(disallowed, func(needle string) bool {
		return strings.Contains(xlsxExecCmd.Long, needle)
	}) {
		t.Fatalf("help text should describe behavior, not endpoint paths")
	}
}

func TestRunExec_RejectsNonPositiveLimits(t *testing.T) {
	resetExecTestGlobals(t)
	filePath, _ := writeWorkbookForExecTest(t)

	cmd := newExecTestCommand()
	if err := cmd.Flags().Set("code", "return 1;"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}
	if err := cmd.Flags().Set("timeout-ms", "0"); err != nil {
		t.Fatalf("setting --timeout-ms: %v", err)
	}
	err := runExec(cmd, []string{filePath})
	if err == nil || !strings.Contains(err.Error(), "--timeout-ms must be > 0") {
		t.Fatalf("unexpected timeout validation error: %v", err)
	}

	cmd = newExecTestCommand()
	if err := cmd.Flags().Set("code", "return 1;"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}
	if err := cmd.Flags().Set("max-output-chars", "-1"); err != nil {
		t.Fatalf("setting --max-output-chars: %v", err)
	}
	err = runExec(cmd, []string{filePath})
	if err == nil || !strings.Contains(err.Error(), "--max-output-chars must be > 0") {
		t.Fatalf("unexpected max-output validation error: %v", err)
	}
}

func TestRunExec_StatelessSuccessHumanOutputAndNoOverwrite(t *testing.T) {
	resetExecTestGlobals(t)
	filePath, originalBytes := writeWorkbookForExecTest(t)

	var gotExecCode string
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/xlsx/exec" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if err := r.ParseMultipartForm(10 << 20); err != nil {
			t.Fatalf("parsing multipart form: %v", err)
		}
		var payload map[string]any
		if err := json.Unmarshal([]byte(r.FormValue("exec")), &payload); err != nil {
			t.Fatalf("parsing exec payload: %v", err)
		}
		gotExecCode, _ = payload["code"].(string)

		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"hello\n","result":{"answer":42}}`)
	}))
	defer server.Close()

	stateless = true
	apiURL = server.URL
	apiKey = "test-key"

	cmd := newExecTestCommand()
	if err := cmd.Flags().Set("code", "return 42;"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}

	output, err := captureExecStdout(t, func() error {
		return runExec(cmd, []string{filePath})
	})
	if err != nil {
		t.Fatalf("runExec failed: %v", err)
	}
	if gotExecCode != "return 42;" {
		t.Fatalf("unexpected exec code sent: %q", gotExecCode)
	}
	if output != "hello\n{\n  \"answer\": 42\n}\n" {
		t.Fatalf("unexpected output:\n%s", output)
	}

	after, err := os.ReadFile(filePath)
	if err != nil {
		t.Fatalf("reading workbook after exec: %v", err)
	}
	if string(after) != string(originalBytes) {
		t.Fatal("workbook bytes changed, but exec must not overwrite local file")
	}
}

func TestRunExec_OkFalseReturnsExit1AndSummary(t *testing.T) {
	resetExecTestGlobals(t)
	filePath, _ := writeWorkbookForExecTest(t)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":false,"stdout":"","error":{"type":"runtime","code":"EXEC_RUNTIME_ERROR","message":"boom"}}`)
	}))
	defer server.Close()

	stateless = true
	apiURL = server.URL
	apiKey = "test-key"

	cmd := newExecTestCommand()
	if err := cmd.Flags().Set("code", "throw new Error('boom')"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}

	output, err := captureExecStdout(t, func() error {
		return runExec(cmd, []string{filePath})
	})
	var exitErr *ExitError
	if err == nil || !errors.As(err, &exitErr) || exitErr.Code != 1 {
		t.Fatalf("expected ExitError code 1, got %v", err)
	}
	if !strings.Contains(output, "runtime (EXEC_RUNTIME_ERROR): boom") {
		t.Fatalf("unexpected output: %q", output)
	}
}

func TestRunExec_StatefulReuploadsOnNotFound(t *testing.T) {
	resetExecTestGlobals(t)
	filePath, _ := writeWorkbookForExecTest(t)

	uploadCalls := 0
	execCalls := 0

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodPost && r.URL.Path == "/v0/files":
			uploadCalls++
			rev := "rev_1"
			if uploadCalls == 2 {
				rev = "rev_2"
			}
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprintf(w, `{"id":"file_1","object":"file","filename":"book.xlsx","bytes":8,"revision_id":"%s","status":"ready"}`, rev)
		case r.Method == http.MethodPost && r.URL.Path == "/v0/files/file_1/xlsx/exec":
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

	cmd := newExecTestCommand()
	if err := cmd.Flags().Set("code", "return true;"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}

	output, err := captureExecStdout(t, func() error {
		return runExec(cmd, []string{filePath})
	})
	if err != nil {
		t.Fatalf("runExec failed: %v", err)
	}
	if uploadCalls != 2 {
		t.Fatalf("expected 2 upload calls, got %d", uploadCalls)
	}
	if execCalls != 2 {
		t.Fatalf("expected 2 files exec calls, got %d", execCalls)
	}
	if output != "done\n{\n  \"ok\": true\n}\n" {
		t.Fatalf("unexpected output:\n%s", output)
	}
}

func TestRunExec_JSONOutputRawEnvelope(t *testing.T) {
	resetExecTestGlobals(t)
	filePath, _ := writeWorkbookForExecTest(t)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"json\n","result":{"a":1},"writes_detected":false}`)
	}))
	defer server.Close()

	stateless = true
	apiURL = server.URL
	apiKey = "test-key"
	jsonOutput = true

	cmd := newExecTestCommand()
	if err := cmd.Flags().Set("code", "return 1"); err != nil {
		t.Fatalf("setting --code: %v", err)
	}

	output, err := captureExecStdout(t, func() error {
		return runExec(cmd, []string{filePath})
	})
	if err != nil {
		t.Fatalf("runExec failed: %v", err)
	}

	var envelope map[string]any
	if err := json.Unmarshal([]byte(output), &envelope); err != nil {
		t.Fatalf("output should be valid JSON, got %q: %v", output, err)
	}
	if envelope["ok"] != true {
		t.Fatalf("unexpected envelope ok: %#v", envelope["ok"])
	}
	if envelope["stdout"] != "json\n" {
		t.Fatalf("unexpected envelope stdout: %#v", envelope["stdout"])
	}
	if _, ok := envelope["result"]; !ok {
		t.Fatalf("result missing from envelope: %#v", envelope)
	}
}

func resetExecTestGlobals(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	origJSONOutput := jsonOutput
	origExecCode := execCode
	origExecScript := execScript
	origExecStdin := execStdin
	origExecExpr := execExpr
	origExecInputJSON := execInputJSON
	origExecTimeoutMS := execTimeoutMS
	origExecMaxOutputChars := execMaxOutputChars

	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
		jsonOutput = origJSONOutput
		execCode = origExecCode
		execScript = origExecScript
		execStdin = origExecStdin
		execExpr = origExecExpr
		execInputJSON = origExecInputJSON
		execTimeoutMS = origExecTimeoutMS
		execMaxOutputChars = origExecMaxOutputChars
	})

	apiKey = ""
	apiURL = ""
	stateless = false
	jsonOutput = false
	execCode = ""
	execScript = ""
	execStdin = false
	execExpr = ""
	execInputJSON = ""
	execTimeoutMS = 0
	execMaxOutputChars = 0
}

func newExecTestCommand() *cobra.Command {
	cmd := &cobra.Command{}
	cmd.Flags().StringVar(&execCode, "code", "", "")
	cmd.Flags().StringVar(&execScript, "script", "", "")
	cmd.Flags().BoolVar(&execStdin, "stdin", false, "")
	cmd.Flags().StringVar(&execExpr, "expr", "", "")
	cmd.Flags().StringVar(&execInputJSON, "input-json", "", "")
	cmd.Flags().IntVar(&execTimeoutMS, "timeout-ms", 0, "")
	cmd.Flags().IntVar(&execMaxOutputChars, "max-output-chars", 0, "")
	return cmd
}

func writeWorkbookForExecTest(t *testing.T) (string, []byte) {
	t.Helper()
	path := filepath.Join(t.TempDir(), "book.xlsx")
	content := []byte{0x50, 0x4b, 0x03, 0x04, 'w', 'i', 't', 'a', 'n'}
	if err := os.WriteFile(path, content, 0o644); err != nil {
		t.Fatalf("writing workbook: %v", err)
	}
	return path, content
}

func captureExecStdout(t *testing.T, fn func() error) (string, error) {
	t.Helper()
	orig := os.Stdout
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatalf("creating stdout pipe: %v", err)
	}
	os.Stdout = w

	runErr := fn()

	if closeErr := w.Close(); closeErr != nil {
		t.Fatalf("closing write pipe: %v", closeErr)
	}
	os.Stdout = orig

	out, readErr := io.ReadAll(r)
	if readErr != nil {
		t.Fatalf("reading captured stdout: %v", readErr)
	}
	if closeErr := r.Close(); closeErr != nil {
		t.Fatalf("closing read pipe: %v", closeErr)
	}
	return string(out), runErr
}
