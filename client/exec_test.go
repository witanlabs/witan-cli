package client

import (
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"
)

func TestExec_PostMultipartRequestShape(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "book.xlsx")
	fileBytes := []byte{0x50, 0x4b, 0x03, 0x04, 'a', 'b', 'c'}
	if err := os.WriteFile(filePath, fileBytes, 0o644); err != nil {
		t.Fatalf("writing temp workbook: %v", err)
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			t.Fatalf("expected POST, got %s", r.Method)
		}
		if r.URL.Path != "/v0/xlsx/exec" {
			t.Fatalf("unexpected path: %s", r.URL.Path)
		}
		if got := r.URL.Query().Get("save"); got != "" {
			t.Fatalf("expected no save query by default, got %q", got)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer test-key" {
			t.Fatalf("unexpected auth header: %q", got)
		}

		if err := r.ParseMultipartForm(10 << 20); err != nil {
			t.Fatalf("parsing multipart form: %v", err)
		}

		f, hdr, err := r.FormFile("file")
		if err != nil {
			t.Fatalf("reading file part: %v", err)
		}
		defer f.Close()
		if hdr.Filename != "book.xlsx" {
			t.Fatalf("unexpected filename: %q", hdr.Filename)
		}
		gotFileBytes, err := io.ReadAll(f)
		if err != nil {
			t.Fatalf("reading file bytes: %v", err)
		}
		if string(gotFileBytes) != string(fileBytes) {
			t.Fatalf("unexpected uploaded file bytes: %q", string(gotFileBytes))
		}

		execField := r.FormValue("exec")
		if execField == "" {
			t.Fatal("expected exec form field")
		}
		var payload map[string]any
		if err := json.Unmarshal([]byte(execField), &payload); err != nil {
			t.Fatalf("decoding exec field: %v", err)
		}
		if payload["code"] != "return input.x;" {
			t.Fatalf("unexpected code: %#v", payload["code"])
		}
		input, ok := payload["input"].(map[string]any)
		if !ok {
			t.Fatalf("input should be object, got %T", payload["input"])
		}
		if input["x"] != float64(7) {
			t.Fatalf("unexpected input.x: %#v", input["x"])
		}
		if payload["timeout_ms"] != float64(2500) {
			t.Fatalf("unexpected timeout_ms: %#v", payload["timeout_ms"])
		}
		if payload["max_output_chars"] != float64(128) {
			t.Fatalf("unexpected max_output_chars: %#v", payload["max_output_chars"])
		}

		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"hello\n","result":{"value":42},"accesses":[{"operation":"read","address":"Sheet1!A1"}]}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", true)
	c.maxAttempts = 1

	resp, err := c.Exec(filePath, ExecRequest{
		Code:           "return input.x;",
		Input:          map[string]any{"x": 7},
		TimeoutMS:      2500,
		MaxOutputChars: 128,
	}, false)
	if err != nil {
		t.Fatalf("Exec failed: %v", err)
	}
	if !resp.Ok {
		t.Fatalf("expected ok response, got %#v", resp)
	}
	if resp.Stdout != "hello\n" {
		t.Fatalf("unexpected stdout: %q", resp.Stdout)
	}
	if string(resp.Result) != `{"value":42}` {
		t.Fatalf("unexpected result: %s", string(resp.Result))
	}
	if len(resp.Accesses) != 1 || resp.Accesses[0].Address != "Sheet1!A1" {
		t.Fatalf("unexpected accesses: %#v", resp.Accesses)
	}
}

func TestExec_ParsesOkFalseEnvelope(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "book.xlsx")
	if err := os.WriteFile(filePath, []byte{0x50, 0x4b, 0x03, 0x04}, 0o644); err != nil {
		t.Fatalf("writing temp workbook: %v", err)
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":false,"stdout":"","error":{"type":"runtime","code":"EXEC_RUNTIME_ERROR","message":"boom"}}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", true)
	c.maxAttempts = 1

	resp, err := c.Exec(filePath, ExecRequest{Code: "throw new Error('boom')"}, false)
	if err != nil {
		t.Fatalf("Exec failed: %v", err)
	}
	if resp.Ok {
		t.Fatalf("expected ok=false envelope, got %#v", resp)
	}
	if resp.Error == nil || resp.Error.Code != "EXEC_RUNTIME_ERROR" {
		t.Fatalf("unexpected error payload: %#v", resp.Error)
	}
}

func TestExec_SaveQueryParam(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "book.xlsx")
	if err := os.WriteFile(filePath, []byte{0x50, 0x4b, 0x03, 0x04}, 0o644); err != nil {
		t.Fatalf("writing temp workbook: %v", err)
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.URL.Query().Get("save"); got != "true" {
			t.Fatalf("expected save=true, got %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"","result":null}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", true)
	c.maxAttempts = 1

	if _, err := c.Exec(filePath, ExecRequest{Code: "return 1"}, true); err != nil {
		t.Fatalf("Exec failed: %v", err)
	}
}

func TestExec_Non200ReturnsAPIError(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "book.xlsx")
	if err := os.WriteFile(filePath, []byte{0x50, 0x4b, 0x03, 0x04}, 0o644); err != nil {
		t.Fatalf("writing temp workbook: %v", err)
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusBadRequest)
		fmt.Fprint(w, `{"error":{"code":"INVALID_ARG","message":"bad request"}}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", true)
	c.maxAttempts = 1

	_, err := c.Exec(filePath, ExecRequest{Code: "return 1"}, false)
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	apiErr, ok := err.(*APIError)
	if !ok {
		t.Fatalf("expected APIError, got %T", err)
	}
	if apiErr.StatusCode != http.StatusBadRequest || apiErr.Code != "INVALID_ARG" {
		t.Fatalf("unexpected APIError: %#v", apiErr)
	}
}

func TestFilesExec_PostJSONWithRevisionAndParsesSuccess(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			t.Fatalf("expected POST, got %s", r.Method)
		}
		if r.URL.Path != "/v0/files/file_123/xlsx/exec" {
			t.Fatalf("unexpected path: %s", r.URL.Path)
		}
		if got := r.URL.Query().Get("revision"); got != "rev_9" {
			t.Fatalf("unexpected revision: %q", got)
		}
		if got := r.URL.Query().Get("save"); got != "" {
			t.Fatalf("expected no save query by default, got %q", got)
		}
		if got := r.Header.Get("Content-Type"); got != "application/json" {
			t.Fatalf("unexpected content type: %q", got)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer test-key" {
			t.Fatalf("unexpected auth header: %q", got)
		}

		body, err := io.ReadAll(r.Body)
		if err != nil {
			t.Fatalf("reading request body: %v", err)
		}
		var payload map[string]any
		if err := json.Unmarshal(body, &payload); err != nil {
			t.Fatalf("unmarshal request body: %v", err)
		}
		if payload["code"] != "return 1;" {
			t.Fatalf("unexpected code: %#v", payload["code"])
		}

		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"done\n","result":{"ok":true}}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", false)
	c.maxAttempts = 1

	resp, err := c.FilesExec("file_123", "rev_9", ExecRequest{Code: "return 1;"}, false)
	if err != nil {
		t.Fatalf("FilesExec failed: %v", err)
	}
	if !resp.Ok || string(resp.Result) != `{"ok":true}` {
		t.Fatalf("unexpected response: %#v", resp)
	}
}

func TestFilesExec_SaveQueryParam(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.URL.Query().Get("revision"); got != "rev_9" {
			t.Fatalf("unexpected revision: %q", got)
		}
		if got := r.URL.Query().Get("save"); got != "true" {
			t.Fatalf("expected save=true, got %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"","result":null}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", false)
	c.maxAttempts = 1

	if _, err := c.FilesExec("file_123", "rev_9", ExecRequest{Code: "return 1;"}, true); err != nil {
		t.Fatalf("FilesExec failed: %v", err)
	}
}

func TestFilesExec_ParsesOkFalseEnvelope(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":false,"stdout":"","error":{"type":"timeout","code":"EXEC_TIMEOUT","message":"timed out"}}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", false)
	c.maxAttempts = 1

	resp, err := c.FilesExec("file_123", "rev_9", ExecRequest{Code: "while(true){}"}, false)
	if err != nil {
		t.Fatalf("FilesExec failed: %v", err)
	}
	if resp.Ok {
		t.Fatalf("expected ok=false envelope, got %#v", resp)
	}
	if resp.Error == nil || resp.Error.Code != "EXEC_TIMEOUT" {
		t.Fatalf("unexpected error payload: %#v", resp.Error)
	}
}

func TestFilesExec_Non200ReturnsAPIError(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusNotFound)
		fmt.Fprint(w, `{"error":{"code":"NOT_FOUND","message":"missing revision"}}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", false)
	c.maxAttempts = 1

	_, err := c.FilesExec("file_123", "rev_9", ExecRequest{Code: "return 1"}, false)
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	apiErr, ok := err.(*APIError)
	if !ok {
		t.Fatalf("expected APIError, got %T", err)
	}
	if apiErr.StatusCode != http.StatusNotFound || apiErr.Code != "NOT_FOUND" {
		t.Fatalf("unexpected APIError: %#v", apiErr)
	}
}
