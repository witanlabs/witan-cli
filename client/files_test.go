package client

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"
)

func TestEnsureUploaded_CacheHitMatchingHashSkipsNetwork(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "test.xlsx")
	if err := os.WriteFile(filePath, []byte("v1"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}
	c.maxAttempts = 1

	hash, err := hashFile(filePath)
	if err != nil {
		t.Fatalf("hashFile: %v", err)
	}
	c.cache.Put(filePath, c.BaseURL, "", CacheEntry{
		FileID: "file_cached", RevisionID: "rev_cached", ContentHash: hash,
	})

	fileID, revID, err := c.EnsureUploaded(filePath)
	if err != nil {
		t.Fatalf("EnsureUploaded failed: %v", err)
	}
	if fileID != "file_cached" || revID != "rev_cached" {
		t.Fatalf("unexpected ids: file=%q rev=%q", fileID, revID)
	}
}

func TestEnsureUploaded_ContentChangedPutsNewRevision(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "test.xlsx")
	if err := os.WriteFile(filePath, []byte("v2"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	putCalls := 0
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodPut && r.URL.Path == "/v0/files/file_known":
			putCalls++
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"id":"file_known","object":"file","filename":"test.xlsx","bytes":2,"revision_id":"rev_new","status":"ready"}`)
		case r.Method == http.MethodPost && r.URL.Path == "/v0/files":
			t.Fatalf("unexpected POST upload")
		default:
			t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
		}
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}
	c.maxAttempts = 1
	// Cache an entry with a stale content hash → forces PUT.
	c.cache.Put(filePath, c.BaseURL, "", CacheEntry{
		FileID: "file_known", RevisionID: "rev_old", ContentHash: "sha256:stale",
	})

	fileID, revID, err := c.EnsureUploaded(filePath)
	if err != nil {
		t.Fatalf("EnsureUploaded failed: %v", err)
	}
	if fileID != "file_known" || revID != "rev_new" {
		t.Fatalf("unexpected ids: file=%q rev=%q", fileID, revID)
	}
	if putCalls != 1 {
		t.Fatalf("expected 1 PUT, got %d", putCalls)
	}

	// Cache should be updated with the new revision and current content hash.
	entry, ok := c.cache.Get(filePath, c.BaseURL, "")
	if !ok {
		t.Fatal("expected cache hit after EnsureUploaded")
	}
	currentHash, _ := hashFile(filePath)
	if entry.RevisionID != "rev_new" || entry.ContentHash != currentHash {
		t.Fatalf("unexpected entry after update: %+v", entry)
	}
}

func TestEnsureUploaded_FallsBackToPostWhenPutNotFound(t *testing.T) {
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

	c := New(server.URL, "test-key", "", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}
	c.maxAttempts = 1
	c.cache.Put(filePath, c.BaseURL, "", CacheEntry{
		FileID: "file_missing", RevisionID: "rev_old", ContentHash: "sha256:stale",
	})

	fileID, revID, err := c.EnsureUploaded(filePath)
	if err != nil {
		t.Fatalf("EnsureUploaded failed: %v", err)
	}
	if fileID != "file_new" || revID != "rev_new" {
		t.Fatalf("unexpected ids: file=%q rev=%q", fileID, revID)
	}
	if putCalls != 1 || postCalls != 1 {
		t.Fatalf("expected 1 PUT + 1 POST, got %d PUT + %d POST", putCalls, postCalls)
	}

	entry, ok := c.cache.Get(filePath, c.BaseURL, "")
	if !ok {
		t.Fatal("expected cache hit")
	}
	if entry.FileID != "file_new" {
		t.Fatalf("expected file_new in cache, got %q", entry.FileID)
	}
}

func TestEnsureUploaded_FreshUploadOnCacheMiss(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "test.xlsx")
	if err := os.WriteFile(filePath, []byte("v4"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	postCalls := 0
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method == http.MethodPost && r.URL.Path == "/v0/files" {
			postCalls++
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprint(w, `{"id":"file_fresh","object":"file","filename":"test.xlsx","bytes":2,"revision_id":"rev_fresh","status":"ready"}`)
			return
		}
		t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}
	c.maxAttempts = 1

	fileID, revID, err := c.EnsureUploaded(filePath)
	if err != nil {
		t.Fatalf("EnsureUploaded failed: %v", err)
	}
	if fileID != "file_fresh" || revID != "rev_fresh" {
		t.Fatalf("unexpected ids: file=%q rev=%q", fileID, revID)
	}
	if postCalls != 1 {
		t.Fatalf("expected 1 POST, got %d", postCalls)
	}

	entry, ok := c.cache.Get(filePath, c.BaseURL, "")
	if !ok {
		t.Fatal("expected cache hit after fresh upload")
	}
	currentHash, _ := hashFile(filePath)
	if entry.FileID != "file_fresh" || entry.ContentHash != currentHash {
		t.Fatalf("unexpected entry: %+v", entry)
	}
}

// Item 2 fix: two distinct files with identical bytes must NOT collapse onto one fileID.
func TestEnsureUploaded_IdenticalContentDistinctPathsGetDistinctFileIDs(t *testing.T) {
	tmpDir := t.TempDir()
	pathA := filepath.Join(tmpDir, "report.xlsx")
	pathB := filepath.Join(tmpDir, "report-backup.xlsx")
	contents := []byte("same bytes")
	if err := os.WriteFile(pathA, contents, 0o644); err != nil {
		t.Fatalf("writing pathA: %v", err)
	}
	if err := os.WriteFile(pathB, contents, 0o644); err != nil {
		t.Fatalf("writing pathB: %v", err)
	}

	postCount := 0
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method == http.MethodPost && r.URL.Path == "/v0/files" {
			postCount++
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprintf(w, `{"id":"file_%d","object":"file","filename":"x.xlsx","bytes":10,"revision_id":"rev_%d","status":"ready"}`, postCount, postCount)
			return
		}
		t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}
	c.maxAttempts = 1

	idA, _, err := c.EnsureUploaded(pathA)
	if err != nil {
		t.Fatalf("EnsureUploaded(A): %v", err)
	}
	idB, _, err := c.EnsureUploaded(pathB)
	if err != nil {
		t.Fatalf("EnsureUploaded(B): %v", err)
	}
	if idA == idB {
		t.Fatalf("expected distinct fileIDs for distinct paths with identical content; both got %q", idA)
	}
	if postCount != 2 {
		t.Fatalf("expected 2 POSTs (one per path), got %d", postCount)
	}
}

func TestReuploadFile_EvictsAndPostsFresh(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "test.xlsx")
	if err := os.WriteFile(filePath, []byte("hello"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	postCount := 0
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method == http.MethodPost && r.URL.Path == "/v0/files" {
			postCount++
			w.Header().Set("Content-Type", "application/json")
			fmt.Fprintf(w, `{"id":"file_after","object":"file","filename":"test.xlsx","bytes":5,"revision_id":"rev_after","status":"ready"}`)
			return
		}
		t.Fatalf("unexpected request: %s %s", r.Method, r.URL.Path)
	}))
	defer server.Close()

	c := New(server.URL, "test-key", "", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}
	c.maxAttempts = 1

	// Pre-populate with a stale-but-hash-matching entry for a now-dead fileID.
	hash, _ := hashFile(filePath)
	c.cache.Put(filePath, c.BaseURL, "", CacheEntry{
		FileID: "file_dead", RevisionID: "rev_dead", ContentHash: hash,
	})

	fileID, revID, err := c.ReuploadFile(filePath)
	if err != nil {
		t.Fatalf("ReuploadFile: %v", err)
	}
	if fileID != "file_after" || revID != "rev_after" {
		t.Fatalf("unexpected ids: file=%q rev=%q", fileID, revID)
	}
	if postCount != 1 {
		t.Fatalf("expected 1 POST after eviction, got %d", postCount)
	}
}

func TestUpdateCachedRevision_StoresEntryByPath(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "calc.xlsx")
	if err := os.WriteFile(filePath, []byte("before"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	c := New("http://localhost:3000", "test-key", "", false)
	c.cache = &FileCache{inMemory: make(map[string]CacheEntry)}

	if err := c.UpdateCachedRevision(filePath, "file_1", "rev_1"); err != nil {
		t.Fatalf("UpdateCachedRevision: %v", err)
	}

	if err := os.WriteFile(filePath, []byte("after"), 0o644); err != nil {
		t.Fatalf("writing updated temp file: %v", err)
	}
	if err := c.UpdateCachedRevision(filePath, "file_1", "rev_2"); err != nil {
		t.Fatalf("UpdateCachedRevision: %v", err)
	}

	entry, ok := c.cache.Get(filePath, c.BaseURL, "")
	if !ok {
		t.Fatal("expected cache hit")
	}
	currentHash, _ := hashFile(filePath)
	if entry.FileID != "file_1" || entry.RevisionID != "rev_2" || entry.ContentHash != currentHash {
		t.Fatalf("unexpected entry: %+v", entry)
	}
}
