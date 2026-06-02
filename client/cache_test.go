package client

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
)

func TestProbeWritable(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "cache-test")
	if !probeWritable(target) {
		t.Fatal("expected probeWritable to succeed on temp dir")
	}
	info, err := os.Stat(target)
	if err != nil || !info.IsDir() {
		t.Fatal("expected directory to exist after probeWritable")
	}
}

func TestProbeWritable_Readonly(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "readonly")
	os.MkdirAll(target, 0o555)
	nested := filepath.Join(target, "nested")
	if probeWritable(nested) {
		t.Skip("filesystem doesn't enforce readonly permissions")
	}
}

func TestFileCache_InMemory(t *testing.T) {
	fc := &FileCache{inMemory: make(map[string]CacheEntry)}

	path := "/tmp/test.xlsx"
	baseURL := "http://localhost:3000"

	if _, ok := fc.Get(path, baseURL, ""); ok {
		t.Fatal("expected cache miss")
	}

	entry := CacheEntry{FileID: "file_1", RevisionID: "rev_1", ContentHash: "sha256:abc", Bytes: 100, Filename: "test.xlsx"}
	fc.Put(path, baseURL, "", entry)

	got, ok := fc.Get(path, baseURL, "")
	if !ok {
		t.Fatal("expected cache hit")
	}
	if got.FileID != "file_1" || got.ContentHash != "sha256:abc" {
		t.Fatalf("unexpected entry: %+v", got)
	}

	fc.Evict(path, baseURL, "")
	if _, ok := fc.Get(path, baseURL, ""); ok {
		t.Fatal("expected cache miss after evict")
	}
}

func TestFileCache_Disk(t *testing.T) {
	dir := filepath.Join(t.TempDir(), "witan-test-cache")
	fc := &FileCache{dir: dir, inMemory: make(map[string]CacheEntry)}
	fc.load()

	path := filepath.Join(t.TempDir(), "data.xlsx")
	baseURL := "http://localhost:3000"
	entry := CacheEntry{FileID: "file_2", RevisionID: "rev_2", ContentHash: "sha256:def", Bytes: 200, Filename: "data.xlsx"}

	fc.Put(path, baseURL, "", entry)

	cachePath := filepath.Join(dir, "cache.json")
	if _, err := os.Stat(cachePath); err != nil {
		t.Fatalf("expected cache.json to exist: %v", err)
	}

	fc2 := &FileCache{dir: dir, inMemory: make(map[string]CacheEntry)}
	fc2.load()

	got, ok := fc2.Get(path, baseURL, "")
	if !ok {
		t.Fatal("expected cache hit after reload")
	}
	if got.FileID != "file_2" || got.RevisionID != "rev_2" || got.ContentHash != "sha256:def" {
		t.Fatalf("unexpected entry: %+v", got)
	}

	fc2.Evict(path, baseURL, "")
	if _, ok := fc2.Get(path, baseURL, ""); ok {
		t.Fatal("expected miss after evict")
	}
}

func TestFileCache_DistinctOrgID(t *testing.T) {
	fc := &FileCache{inMemory: make(map[string]CacheEntry)}
	path := "/tmp/test.xlsx"
	baseURL := "http://localhost:3000"

	fc.Put(path, baseURL, "org_aaa", CacheEntry{FileID: "file_a"})
	fc.Put(path, baseURL, "org_bbb", CacheEntry{FileID: "file_b"})

	a, _ := fc.Get(path, baseURL, "org_aaa")
	b, _ := fc.Get(path, baseURL, "org_bbb")
	if a.FileID == b.FileID {
		t.Fatalf("expected distinct entries per orgID, got both %q", a.FileID)
	}
}

func TestFileCache_DistinctBaseURL(t *testing.T) {
	fc := &FileCache{inMemory: make(map[string]CacheEntry)}
	path := "/tmp/test.xlsx"

	fc.Put(path, "http://localhost:3000", "", CacheEntry{FileID: "file_local"})
	fc.Put(path, "https://api.witanlabs.com", "", CacheEntry{FileID: "file_prod"})

	a, _ := fc.Get(path, "http://localhost:3000", "")
	b, _ := fc.Get(path, "https://api.witanlabs.com", "")
	if a.FileID == b.FileID {
		t.Fatalf("expected distinct entries per baseURL, got both %q", a.FileID)
	}
}

func TestFileCache_DistinctPaths(t *testing.T) {
	fc := &FileCache{inMemory: make(map[string]CacheEntry)}
	baseURL := "http://localhost:3000"

	fc.Put("/tmp/report.xlsx", baseURL, "", CacheEntry{FileID: "file_a"})
	fc.Put("/tmp/report-backup.xlsx", baseURL, "", CacheEntry{FileID: "file_b"})

	a, _ := fc.Get("/tmp/report.xlsx", baseURL, "")
	b, _ := fc.Get("/tmp/report-backup.xlsx", baseURL, "")
	if a.FileID == b.FileID {
		t.Fatalf("expected distinct entries per path, got both %q", a.FileID)
	}
}

func TestFileCache_DiscardsOldVersion(t *testing.T) {
	dir := filepath.Join(t.TempDir(), "witan-test-cache-old")
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatalf("mkdir: %v", err)
	}
	// Write a v2 cache.json with entries that should be ignored.
	v2 := []byte(`{"v":2,"files":{"sha256:abc@http://localhost:3000@":{"file_id":"old_file","revision_id":"old_rev"}},"known":{}}`)
	if err := os.WriteFile(filepath.Join(dir, "cache.json"), v2, 0o644); err != nil {
		t.Fatalf("write v2: %v", err)
	}

	fc := &FileCache{dir: dir, inMemory: make(map[string]CacheEntry)}
	fc.load()

	if fc.data.Version != cacheVersion {
		t.Fatalf("expected version %d after load, got %d", cacheVersion, fc.data.Version)
	}
	if len(fc.data.Entries) != 0 {
		t.Fatalf("expected empty entries after discarding v2, got %d", len(fc.data.Entries))
	}
}

func TestFileCache_PersistedJSONShape(t *testing.T) {
	dir := filepath.Join(t.TempDir(), "witan-test-cache-shape")
	fc := &FileCache{dir: dir, inMemory: make(map[string]CacheEntry)}
	fc.load()

	fc.Put("/tmp/x.xlsx", "http://localhost:3000", "org_z", CacheEntry{
		FileID: "file_x", RevisionID: "rev_x", ContentHash: "sha256:xx", Bytes: 7, Filename: "x.xlsx",
	})

	raw, err := os.ReadFile(filepath.Join(dir, "cache.json"))
	if err != nil {
		t.Fatalf("read cache.json: %v", err)
	}
	var on cacheData
	if err := json.Unmarshal(raw, &on); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if on.Version != cacheVersion {
		t.Fatalf("expected v%d, got v%d", cacheVersion, on.Version)
	}
	if len(on.Entries) != 1 {
		t.Fatalf("expected 1 entry, got %d", len(on.Entries))
	}
}

func TestEntryKey_PathInKey(t *testing.T) {
	a := entryKey("/tmp/a.xlsx", "http://localhost:3000", "")
	b := entryKey("/tmp/b.xlsx", "http://localhost:3000", "")
	if a == b {
		t.Fatal("expected distinct keys for distinct paths")
	}
}

func TestEntryKey_RelativePathResolvesToAbsolute(t *testing.T) {
	cwd, err := os.Getwd()
	if err != nil {
		t.Fatalf("getwd: %v", err)
	}
	rel := entryKey("./foo.xlsx", "http://localhost:3000", "")
	abs := entryKey(filepath.Join(cwd, "foo.xlsx"), "http://localhost:3000", "")
	if rel != abs {
		t.Fatalf("expected relative path to resolve to absolute; got %q vs %q", rel, abs)
	}
}

func TestHashFile_ContentDistinguishes(t *testing.T) {
	dir := t.TempDir()
	path1 := filepath.Join(dir, "a.xlsx")
	path2 := filepath.Join(dir, "b.xlsx")
	os.WriteFile(path1, []byte("hello"), 0o644)
	os.WriteFile(path2, []byte("world"), 0o644)

	h1, err := hashFile(path1)
	if err != nil {
		t.Fatalf("hashFile: %v", err)
	}
	h2, err := hashFile(path2)
	if err != nil {
		t.Fatalf("hashFile: %v", err)
	}
	if h1 == h2 {
		t.Fatal("expected different hashes for different content")
	}

	// Same content → same hash
	os.WriteFile(path2, []byte("hello"), 0o644)
	h3, _ := hashFile(path2)
	if h1 != h3 {
		t.Fatal("expected same hash for identical content")
	}
}

func TestIsNotFound(t *testing.T) {
	err404 := &APIError{StatusCode: 404, Code: "not_found", Message: "file not found"}
	if !IsNotFound(err404) {
		t.Fatal("expected IsNotFound to be true for 404")
	}

	route404 := &APIError{StatusCode: 404, Code: "not_found", Message: "Route POST /v0/orgs/org_1/pptx/exec not found"}
	if IsNotFound(route404) {
		t.Fatal("expected IsNotFound to be false for unmounted route 404")
	}

	err500 := &APIError{StatusCode: 500, Code: "internal", Message: "server error"}
	if IsNotFound(err500) {
		t.Fatal("expected IsNotFound to be false for 500")
	}

	if IsNotFound(nil) {
		t.Fatal("expected IsNotFound to be false for nil")
	}
}
