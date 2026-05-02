package cmd

import (
	"bytes"
	"context"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"sync/atomic"
	"testing"

	"github.com/coder/websocket"
	"github.com/witanlabs/witan-cli/client"
)

func TestRPCStatelessSaveWritesFileAndRedactsMeta(t *testing.T) {
	filePath := filepath.Join(t.TempDir(), "book.xlsx")
	if err := os.WriteFile(filePath, []byte("before"), 0o644); err != nil {
		t.Fatalf("writing workbook: %v", err)
	}

	s := &rpcSession{mode: "stateless", filePath: filePath}
	raw := fmt.Sprintf(
		`{"id":"save-1","ok":true,"result":true,"meta":{"file":%q,"content_type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}}`,
		base64.StdEncoding.EncodeToString([]byte("after")),
	)

	redacted, err := s.applyRPCResponseSideEffects(
		rpcRequestEnvelope{ID: "save-1", Op: "save"},
		[]byte(raw),
	)
	if err != nil {
		t.Fatalf("applyRPCResponseSideEffects failed: %v", err)
	}

	got, err := os.ReadFile(filePath)
	if err != nil {
		t.Fatalf("reading workbook: %v", err)
	}
	if string(got) != "after" {
		t.Fatalf("expected workbook writeback, got %q", got)
	}
	assertNoRPCMeta(t, redacted)
}

func TestRPCFilesSaveDownloadsRevisionUpdatesCacheAndRedactsMeta(t *testing.T) {
	t.Setenv("TMPDIR", t.TempDir())
	filePath := filepath.Join(t.TempDir(), "book.xlsx")
	if err := os.WriteFile(filePath, []byte("before"), 0o644); err != nil {
		t.Fatalf("writing workbook: %v", err)
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet || r.URL.Path != "/v0/files/file_1/content" {
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.String())
		}
		if got := r.URL.Query().Get("revision"); got != "rev_2" {
			t.Fatalf("unexpected revision: %q", got)
		}
		fmt.Fprint(w, "after")
	}))
	defer server.Close()

	c := client.New(server.URL, "key", "", false)
	s := &rpcSession{mode: "files", client: c, filePath: filePath, fileID: "file_1"}
	raw := []byte(`{"id":"save-1","ok":true,"result":true,"meta":{"revision_id":"rev_2"}}`)

	redacted, err := s.applyRPCResponseSideEffects(
		rpcRequestEnvelope{ID: "save-1", Op: "save"},
		raw,
	)
	if err != nil {
		t.Fatalf("applyRPCResponseSideEffects failed: %v", err)
	}

	got, err := os.ReadFile(filePath)
	if err != nil {
		t.Fatalf("reading workbook: %v", err)
	}
	if string(got) != "after" {
		t.Fatalf("expected workbook writeback, got %q", got)
	}
	assertNoRPCMeta(t, redacted)

	fileID, revisionID, err := c.EnsureUploaded(filePath)
	if err != nil {
		t.Fatalf("EnsureUploaded after save failed: %v", err)
	}
	if fileID != "file_1" || revisionID != "rev_2" {
		t.Fatalf("cache not updated, got file=%q revision=%q", fileID, revisionID)
	}
}

func TestRPCFilesRetriesStaleSessionWithReupload(t *testing.T) {
	filePath := filepath.Join(t.TempDir(), "book.xlsx")
	if err := os.WriteFile(filePath, []byte("workbook"), 0o644); err != nil {
		t.Fatalf("writing workbook: %v", err)
	}

	var staleWSRequests atomic.Int32
	var freshWSRequests atomic.Int32
	var uploadRequests atomic.Int32

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodGet && r.URL.Path == "/v0/files/file_1/xlsx/ws":
			staleWSRequests.Add(1)
			serveRPCWebSocket(t, w, r, func(ctx context.Context, conn *websocket.Conn) {
				_, _, err := conn.Read(ctx)
				if err != nil {
					t.Errorf("reading stale ws request: %v", err)
					return
				}
				if err := conn.Write(ctx, websocket.MessageText, []byte(`{"op":"error","code":"FILE_NOT_FOUND","message":"missing"}`)); err != nil {
					t.Errorf("writing stale ws response: %v", err)
				}
				_ = conn.Close(4404, "missing")
			})
		case r.Method == http.MethodPost && r.URL.Path == "/v0/files":
			uploadRequests.Add(1)
			fmt.Fprint(w, `{"id":"file_2","object":"file","filename":"book.xlsx","bytes":8,"revision_id":"rev_2","status":"ready"}`)
		case r.Method == http.MethodGet && r.URL.Path == "/v0/files/file_2/xlsx/ws":
			freshWSRequests.Add(1)
			if got := r.URL.Query().Get("revision"); got != "rev_2" {
				t.Errorf("unexpected fresh revision: %q", got)
			}
			if got := r.URL.Query().Get("protocol"); got != "rpc" {
				t.Errorf("unexpected fresh protocol: %q", got)
			}
			if got := r.URL.Query().Get("hint"); got != "Sheet1" {
				t.Errorf("unexpected fresh hint: %q", got)
			}
			if got := r.URL.Query().Get("locale"); got != "fr-FR" {
				t.Errorf("unexpected fresh locale: %q", got)
			}
			serveRPCWebSocket(t, w, r, func(ctx context.Context, conn *websocket.Conn) {
				_, raw, err := conn.Read(ctx)
				if err != nil {
					t.Errorf("reading fresh ws request: %v", err)
					return
				}
				if string(raw) != `{"id":"1","op":"listSheets","args":{}}` {
					t.Errorf("unexpected replayed request: %s", raw)
				}
				if err := conn.Write(ctx, websocket.MessageText, []byte(`{"id":"1","ok":true,"result":{"ready":true}}`)); err != nil {
					t.Errorf("writing fresh ws response: %v", err)
				}
			})
		default:
			t.Errorf("unexpected request: %s %s", r.Method, r.URL.String())
			http.NotFound(w, r)
		}
	}))
	defer server.Close()

	c := client.New(server.URL, "", "", false)
	staleURL, err := c.FilesXlsxRPCWebSocketURL("file_1", "rev_1", "Sheet1", "fr-FR")
	if err != nil {
		t.Fatalf("building stale ws URL: %v", err)
	}
	conn, _, err := websocket.Dial(context.Background(), staleURL, nil)
	if err != nil {
		t.Fatalf("dialing stale ws: %v", err)
	}
	conn.SetReadLimit(rpcReadLimit)

	session := &rpcSession{
		conn:     conn,
		client:   c,
		filePath: filePath,
		fileID:   "file_1",
		hint:     "Sheet1",
		locale:   "fr-FR",
		mode:     "files",
	}
	var out bytes.Buffer
	err = relayRPCStdio(
		context.Background(),
		session,
		strings.NewReader(`{"id":"1","op":"listSheets","args":{}}`+"\n"),
		&out,
	)
	if err != nil {
		t.Fatalf("relayRPCStdio failed: %v", err)
	}

	var got map[string]any
	if err := json.Unmarshal(bytes.TrimSpace(out.Bytes()), &got); err != nil {
		t.Fatalf("parsing stdout: %v", err)
	}
	if got["id"] != "1" || got["ok"] != true {
		t.Fatalf("unexpected stdout: %s", out.String())
	}
	if staleWSRequests.Load() != 1 || uploadRequests.Load() != 1 || freshWSRequests.Load() != 1 {
		t.Fatalf("unexpected request counts: stale=%d uploads=%d fresh=%d", staleWSRequests.Load(), uploadRequests.Load(), freshWSRequests.Load())
	}
	if session.fileID != "file_2" {
		t.Fatalf("expected reconnected fileID file_2, got %q", session.fileID)
	}
}

func serveRPCWebSocket(t *testing.T, w http.ResponseWriter, r *http.Request, handle func(context.Context, *websocket.Conn)) {
	t.Helper()
	conn, err := websocket.Accept(w, r, nil)
	if err != nil {
		t.Errorf("accepting websocket: %v", err)
		return
	}
	defer conn.CloseNow()
	handle(context.Background(), conn)
}

func assertNoRPCMeta(t *testing.T, raw []byte) {
	t.Helper()
	var obj map[string]json.RawMessage
	if err := json.Unmarshal(raw, &obj); err != nil {
		t.Fatalf("redacted response is not JSON: %v", err)
	}
	if _, ok := obj["meta"]; ok {
		t.Fatalf("expected meta to be redacted from %s", raw)
	}
}
