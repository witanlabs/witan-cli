package client

import (
	"encoding/json"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestPPTXRender_PostsPPTXRender(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "deck.pptx")
	if err := os.WriteFile(filePath, []byte("pptx bytes"), 0o644); err != nil {
		t.Fatalf("writing presentation: %v", err)
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/orgs/org_1/pptx/render" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.URL.Query().Get("slide"); got != "2" {
			t.Fatalf("expected slide=2, got %q", got)
		}
		if got := r.URL.Query().Get("dpr"); got != "2" {
			t.Fatalf("expected dpr=2, got %q", got)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer test-key" {
			t.Fatalf("unexpected auth header: %q", got)
		}
		if got := r.Header.Get("Content-Type"); got != "application/vnd.openxmlformats-officedocument.presentationml.presentation" {
			t.Fatalf("unexpected content type: %q", got)
		}
		body, err := io.ReadAll(r.Body)
		if err != nil {
			t.Fatalf("reading body: %v", err)
		}
		if string(body) != "pptx bytes" {
			t.Fatalf("unexpected body: %q", string(body))
		}
		w.Header().Set("Content-Type", "image/png")
		fmt.Fprint(w, "png bytes")
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "org_1", true)
	c.maxAttempts = 1

	body, contentType, err := c.PPTXRender(filePath, map[string]string{"slide": "2", "dpr": "2"})
	if err != nil {
		t.Fatalf("PPTXRender failed: %v", err)
	}
	if string(body) != "png bytes" {
		t.Fatalf("unexpected response body: %q", string(body))
	}
	if contentType != "image/png" {
		t.Fatalf("unexpected response content type: %q", contentType)
	}
}

func TestPPTXExec_PostsMultipartPPTXExec(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "deck.pptx")
	if err := os.WriteFile(filePath, []byte("pptx bytes"), 0o644); err != nil {
		t.Fatalf("writing presentation: %v", err)
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/orgs/org_1/pptx/exec" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.URL.Query().Get("save"); got != "true" {
			t.Fatalf("expected save=true, got %q", got)
		}
		if got := r.URL.Query().Get("locale"); got != "en-US" {
			t.Fatalf("expected locale=en-US, got %q", got)
		}
		if got := r.Header.Get("Accept-Language"); got != "en-US" {
			t.Fatalf("unexpected Accept-Language: %q", got)
		}

		form := parsePPTXMultipartExecForm(t, r)
		assertPPTXMultipartFile(t, form, "file", "pptx bytes")
		var execReq ExecRequest
		if err := json.Unmarshal([]byte(form.Value["exec"][0]), &execReq); err != nil {
			t.Fatalf("unmarshal exec field: %v", err)
		}
		if execReq.Code != "return 1;" || execReq.Locale != "en-US" || execReq.TimeoutMS != 5000 || execReq.MaxOutputChars != 123 {
			t.Fatalf("unexpected exec request: %+v", execReq)
		}
		input, ok := execReq.Input.(map[string]any)
		if !ok || input["name"] != "deck" {
			t.Fatalf("unexpected input: %#v", execReq.Input)
		}

		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"stdout":"hello\n","result":{"count":1}}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "org_1", true)
	c.maxAttempts = 1

	result, err := c.PPTXExec(filePath, ExecRequest{
		Code:           "return 1;",
		Input:          map[string]any{"name": "deck"},
		Locale:         "en-US",
		TimeoutMS:      5000,
		MaxOutputChars: 123,
	}, true)
	if err != nil {
		t.Fatalf("PPTXExec failed: %v", err)
	}
	if !result.Ok || result.Stdout != "hello\n" || string(result.Result) != `{"count":1}` {
		t.Fatalf("unexpected result: %+v", result)
	}
}

func TestPPTXExecCreate_PostsCreateQueryAndFilenameWithoutFile(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/orgs/org_1/pptx/exec" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		if got := r.URL.Query().Get("create"); got != "true" {
			t.Fatalf("expected create=true, got %q", got)
		}
		if got := r.URL.Query().Get("save"); got != "true" {
			t.Fatalf("expected save=true, got %q", got)
		}

		form := parsePPTXMultipartExecForm(t, r)
		if _, ok := form.File["file"]; ok {
			t.Fatalf("did not expect file part for create-mode exec")
		}
		var execReq ExecRequest
		if err := json.Unmarshal([]byte(form.Value["exec"][0]), &execReq); err != nil {
			t.Fatalf("unmarshal exec field: %v", err)
		}
		if execReq.Code != "return true;" {
			t.Fatalf("unexpected code: %q", execReq.Code)
		}
		if execReq.Filename != "new.pptx" {
			t.Fatalf("unexpected filename: %q", execReq.Filename)
		}

		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"file":"cHB0eA=="}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "org_1", true)
	c.maxAttempts = 1

	result, err := c.PPTXExecCreate(filepath.Join(t.TempDir(), "new.pptx"), ExecRequest{Code: "return true;"}, true)
	if err != nil {
		t.Fatalf("PPTXExecCreate failed: %v", err)
	}
	if result.File == nil || *result.File != "cHB0eA==" {
		t.Fatalf("unexpected result: %+v", result)
	}
}

func TestParsePPTXAPIError_InvalidMIMETypeMessage(t *testing.T) {
	err := parsePPTXAPIError(http.StatusBadRequest, []byte(`{"error":{"code":"invalid_mime_type","message":"Unsupported Content-Type: text/plain"}}`), "")
	apiErr, ok := err.(*APIError)
	if !ok {
		t.Fatalf("expected APIError, got %T", err)
	}
	if got := apiErr.Error(); got != "unsupported file type - expected .pptx" {
		t.Fatalf("unexpected invalid MIME message: %q", got)
	}
}

func TestFilesPPTXExec_PostsPPTXExecJSON(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v0/orgs/org_1/files/file_1/pptx/exec" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		q := r.URL.Query()
		if q.Get("revision") != "rev_1" || q.Get("save") != "true" || q.Get("locale") != "pt-PT" {
			t.Fatalf("unexpected query: %s", r.URL.RawQuery)
		}
		if got := r.Header.Get("Content-Type"); !strings.HasPrefix(got, "application/json") {
			t.Fatalf("unexpected content type: %q", got)
		}
		var execReq ExecRequest
		if err := json.NewDecoder(r.Body).Decode(&execReq); err != nil {
			t.Fatalf("decode body: %v", err)
		}
		if execReq.Code != "return 1;" || execReq.Locale != "pt-PT" {
			t.Fatalf("unexpected exec request: %+v", execReq)
		}

		w.Header().Set("Content-Type", "application/json")
		fmt.Fprint(w, `{"ok":true,"revision_id":"rev_2"}`)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "org_1", false)
	c.maxAttempts = 1

	result, err := c.FilesPPTXExec("file_1", "rev_1", ExecRequest{Code: "return 1;", Locale: "pt-PT"}, true)
	if err != nil {
		t.Fatalf("FilesPPTXExec failed: %v", err)
	}
	if result.RevisionID == nil || *result.RevisionID != "rev_2" {
		t.Fatalf("unexpected result: %+v", result)
	}
}

func TestFilesPPTXRender_GetsPPTXRender(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet || r.URL.Path != "/v0/orgs/org_1/files/file_1/pptx/render" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
		q := r.URL.Query()
		if q.Get("revision") != "rev_1" || q.Get("slide") != "3" || q.Get("dpr") != "1" {
			t.Fatalf("unexpected query: %s", r.URL.RawQuery)
		}
		w.Header().Set("Content-Type", "image/png")
		fmt.Fprint(w, "png bytes")
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "org_1", false)
	c.maxAttempts = 1

	body, contentType, err := c.FilesPPTXRender("file_1", "rev_1", map[string]string{"slide": "3", "dpr": "1"})
	if err != nil {
		t.Fatalf("FilesPPTXRender failed: %v", err)
	}
	if string(body) != "png bytes" || contentType != "image/png" {
		t.Fatalf("unexpected response: body=%q contentType=%q", string(body), contentType)
	}
}

func parsePPTXMultipartExecForm(t *testing.T, r *http.Request) *multipart.Form {
	t.Helper()
	if err := r.ParseMultipartForm(1 << 20); err != nil {
		t.Fatalf("ParseMultipartForm: %v", err)
	}
	if len(r.MultipartForm.Value["exec"]) != 1 {
		t.Fatalf("expected one exec field, got %d", len(r.MultipartForm.Value["exec"]))
	}
	return r.MultipartForm
}

func assertPPTXMultipartFile(t *testing.T, form *multipart.Form, name, want string) {
	t.Helper()
	files := form.File[name]
	if len(files) != 1 {
		t.Fatalf("expected one %q file part, got %d", name, len(files))
	}
	f, err := files[0].Open()
	if err != nil {
		t.Fatalf("opening %q file part: %v", name, err)
	}
	defer f.Close()
	body, err := io.ReadAll(f)
	if err != nil {
		t.Fatalf("reading %q file part: %v", name, err)
	}
	if string(body) != want {
		t.Fatalf("unexpected %q file body: %q", name, string(body))
	}
}
