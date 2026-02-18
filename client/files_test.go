package client

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"
)

func TestEnsureUploaded_UsesKnownFileVersionUpload(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "test.xlsx")
	if err := os.WriteFile(filePath, []byte("v2"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	putCalls := 0
	postCalls := 0
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodPut && r.URL.Path == "/v0/files/file_known":
			putCalls++
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"id":"file_known","object":"file","filename":"test.xlsx","bytes":2,"revision_id":"rev_new","status":"ready"}`)
		case r.Method == http.MethodPost && r.URL.Path == "/v0/files":
			postCalls++
			t.Fatalf("unexpected POST upload")
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer server.Close()

	c := New(server.URL, "test-key", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}
	c.maxAttempts = 1
	c.cache.PutKnown(filePath, c.BaseURL, CacheEntry{
		FileID:     "file_known",
		RevisionID: "rev_old",
		Filename:   "test.xlsx",
	})

	fileID, revID, err := c.EnsureUploaded(filePath)
	if err != nil {
		t.Fatalf("EnsureUploaded failed: %v", err)
	}
	if fileID != "file_known" || revID != "rev_new" {
		t.Fatalf("unexpected ids: file=%q rev=%q", fileID, revID)
	}
	if putCalls != 1 {
		t.Fatalf("expected 1 PUT call, got %d", putCalls)
	}
	if postCalls != 0 {
		t.Fatalf("expected 0 POST calls, got %d", postCalls)
	}

	hashKey, err := HashFile(filePath, c.BaseURL)
	if err != nil {
		t.Fatalf("HashFile failed: %v", err)
	}
	hashEntry, ok := c.cache.Get(hashKey)
	if !ok {
		t.Fatal("expected hash cache hit")
	}
	if hashEntry.RevisionID != "rev_new" {
		t.Fatalf("expected hash revision rev_new, got %q", hashEntry.RevisionID)
	}

	knownEntry, ok := c.cache.GetKnown(filePath, c.BaseURL)
	if !ok {
		t.Fatal("expected known cache hit")
	}
	if knownEntry.RevisionID != "rev_new" {
		t.Fatalf("expected known revision rev_new, got %q", knownEntry.RevisionID)
	}
}

func TestEnsureUploaded_FallsBackToPostWhenKnownFileNotFound(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "test.xlsx")
	if err := os.WriteFile(filePath, []byte("v3"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	putCalls := 0
	postCalls := 0
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodPut && r.URL.Path == "/v0/files/file_missing":
			putCalls++
			w.Header().Set("Content-Type", "application/json")
			w.WriteHeader(http.StatusNotFound)
			fmt.Fprint(w, `{"error":{"code":"file_not_found","message":"File not found"}}`)
		case r.Method == http.MethodPost && r.URL.Path == "/v0/files":
			postCalls++
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"id":"file_new","object":"file","filename":"test.xlsx","bytes":2,"revision_id":"rev_new","status":"ready"}`)
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer server.Close()

	c := New(server.URL, "test-key", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}
	c.maxAttempts = 1
	c.cache.PutKnown(filePath, c.BaseURL, CacheEntry{
		FileID:     "file_missing",
		RevisionID: "rev_old",
		Filename:   "test.xlsx",
	})

	fileID, revID, err := c.EnsureUploaded(filePath)
	if err != nil {
		t.Fatalf("EnsureUploaded failed: %v", err)
	}
	if fileID != "file_new" || revID != "rev_new" {
		t.Fatalf("unexpected ids: file=%q rev=%q", fileID, revID)
	}
	if putCalls != 1 {
		t.Fatalf("expected 1 PUT call, got %d", putCalls)
	}
	if postCalls != 1 {
		t.Fatalf("expected 1 POST call, got %d", postCalls)
	}

	knownEntry, ok := c.cache.GetKnown(filePath, c.BaseURL)
	if !ok {
		t.Fatal("expected known cache hit")
	}
	if knownEntry.FileID != "file_new" || knownEntry.RevisionID != "rev_new" {
		t.Fatalf("unexpected known entry: %+v", knownEntry)
	}
}

func TestUpdateCachedRevision_UpdatesHashAndKnownEntry(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "calc.xlsx")
	if err := os.WriteFile(filePath, []byte("before"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	c := New("http://localhost:3000", "test-key", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}

	if err := c.UpdateCachedRevision(filePath, "file_1", "rev_1"); err != nil {
		t.Fatalf("UpdateCachedRevision failed: %v", err)
	}

	if err := os.WriteFile(filePath, []byte("after"), 0o644); err != nil {
		t.Fatalf("writing updated temp file: %v", err)
	}
	if err := c.UpdateCachedRevision(filePath, "file_1", "rev_2"); err != nil {
		t.Fatalf("UpdateCachedRevision failed: %v", err)
	}

	newHash, err := HashFile(filePath, c.BaseURL)
	if err != nil {
		t.Fatalf("HashFile failed: %v", err)
	}
	hashEntry, ok := c.cache.Get(newHash)
	if !ok {
		t.Fatal("expected hash cache hit")
	}
	if hashEntry.RevisionID != "rev_2" {
		t.Fatalf("expected hash revision rev_2, got %q", hashEntry.RevisionID)
	}

	knownEntry, ok := c.cache.GetKnown(filePath, c.BaseURL)
	if !ok {
		t.Fatal("expected known cache hit")
	}
	if knownEntry.FileID != "file_1" || knownEntry.RevisionID != "rev_2" {
		t.Fatalf("unexpected known entry: %+v", knownEntry)
	}
}
