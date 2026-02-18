package client

import (
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
	// Directory should have been created
	info, err := os.Stat(target)
	if err != nil || !info.IsDir() {
		t.Fatal("expected directory to exist after probeWritable")
	}
}

func TestProbeWritable_Readonly(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "readonly")
	os.MkdirAll(target, 0o555)
	// Can't write a file inside a 555 dir (on most systems)
	nested := filepath.Join(target, "nested")
	if probeWritable(nested) {
		// On some systems (e.g. running as root), 555 doesn't prevent writes
		t.Skip("filesystem doesn't enforce readonly permissions")
	}
}

func TestFileCache_InMemory(t *testing.T) {
	fc := &FileCache{
		inMemory: make(map[string]CacheEntry),
	}

	key := "sha256:abc123@http://localhost"

	// Miss
	_, ok := fc.Get(key)
	if ok {
		t.Fatal("expected cache miss")
	}

	// Put
	entry := CacheEntry{FileID: "file_1", RevisionID: "rev_1", Bytes: 100, Filename: "test.xlsx"}
	fc.Put(key, entry)

	// Hit
	got, ok := fc.Get(key)
	if !ok {
		t.Fatal("expected cache hit")
	}
	if got.FileID != "file_1" {
		t.Fatalf("expected file_1, got %s", got.FileID)
	}

	// Evict
	fc.Evict(key)
	_, ok = fc.Get(key)
	if ok {
		t.Fatal("expected cache miss after evict")
	}
}

func TestFileCache_Disk(t *testing.T) {
	dir := filepath.Join(t.TempDir(), "witan-test-cache")

	fc := &FileCache{dir: dir, inMemory: make(map[string]CacheEntry)}
	fc.load()

	key := "sha256:deadbeef@http://localhost"
	entry := CacheEntry{FileID: "file_2", RevisionID: "rev_2", Bytes: 200, Filename: "data.xlsx"}

	fc.Put(key, entry)

	// Verify the file was written
	cachePath := filepath.Join(dir, "cache.json")
	if _, err := os.Stat(cachePath); err != nil {
		t.Fatalf("expected cache.json to exist: %v", err)
	}

	// Load a fresh cache from the same directory
	fc2 := &FileCache{dir: dir, inMemory: make(map[string]CacheEntry)}
	fc2.load()

	got, ok := fc2.Get(key)
	if !ok {
		t.Fatal("expected cache hit after reload")
	}
	if got.FileID != "file_2" || got.RevisionID != "rev_2" {
		t.Fatalf("unexpected entry: %+v", got)
	}

	// Evict and verify
	fc2.Evict(key)
	_, ok = fc2.Get(key)
	if ok {
		t.Fatal("expected miss after evict")
	}
}

func TestFileCache_KnownInMemory(t *testing.T) {
	fc := &FileCache{
		inMemory: make(map[string]CacheEntry),
	}

	path := "/tmp/test.xlsx"
	baseURL := "http://localhost:3000"
	entry := CacheEntry{FileID: "file_known", RevisionID: "rev_known", Filename: "test.xlsx"}

	_, ok := fc.GetKnown(path, baseURL)
	if ok {
		t.Fatal("expected known miss")
	}

	fc.PutKnown(path, baseURL, entry)
	got, ok := fc.GetKnown(path, baseURL)
	if !ok {
		t.Fatal("expected known hit")
	}
	if got.FileID != "file_known" || got.RevisionID != "rev_known" {
		t.Fatalf("unexpected known entry: %+v", got)
	}

	fc.EvictKnown(path, baseURL)
	_, ok = fc.GetKnown(path, baseURL)
	if ok {
		t.Fatal("expected known miss after evict")
	}
}

func TestFileCache_KnownDisk(t *testing.T) {
	dir := filepath.Join(t.TempDir(), "witan-test-cache-known")

	fc := &FileCache{dir: dir, inMemory: make(map[string]CacheEntry)}
	fc.load()

	path := filepath.Join(t.TempDir(), "book.xlsx")
	baseURL := "http://localhost:3000"
	entry := CacheEntry{FileID: "file_known_2", RevisionID: "rev_known_2", Filename: "book.xlsx"}

	fc.PutKnown(path, baseURL, entry)

	fc2 := &FileCache{dir: dir, inMemory: make(map[string]CacheEntry)}
	fc2.load()

	got, ok := fc2.GetKnown(path, baseURL)
	if !ok {
		t.Fatal("expected known hit after reload")
	}
	if got.FileID != "file_known_2" || got.RevisionID != "rev_known_2" {
		t.Fatalf("unexpected known entry: %+v", got)
	}
}

func TestHashFile(t *testing.T) {
	// Create a temp file to hash
	dir := t.TempDir()
	path := filepath.Join(dir, "test.xlsx")
	os.WriteFile(path, []byte("hello world"), 0o644)

	key1, err := HashFile(path, "http://localhost:3000")
	if err != nil {
		t.Fatalf("HashFile failed: %v", err)
	}

	// Same file, same base URL → same key
	key2, err := HashFile(path, "http://localhost:3000")
	if err != nil {
		t.Fatalf("HashFile failed: %v", err)
	}
	if key1 != key2 {
		t.Fatalf("expected same hash, got %s and %s", key1, key2)
	}

	// Same file, different base URL → different key
	key3, err := HashFile(path, "https://api.witanlabs.com")
	if err != nil {
		t.Fatalf("HashFile failed: %v", err)
	}
	if key1 == key3 {
		t.Fatal("expected different keys for different base URLs")
	}

	// Different content → different key
	path2 := filepath.Join(dir, "test2.xlsx")
	os.WriteFile(path2, []byte("goodbye world"), 0o644)
	key4, err := HashFile(path2, "http://localhost:3000")
	if err != nil {
		t.Fatalf("HashFile failed: %v", err)
	}
	if key1 == key4 {
		t.Fatal("expected different keys for different content")
	}
}

func TestIsNotFound(t *testing.T) {
	err404 := &APIError{StatusCode: 404, Code: "not_found", Message: "file not found"}
	if !IsNotFound(err404) {
		t.Fatal("expected IsNotFound to be true for 404")
	}

	err500 := &APIError{StatusCode: 500, Code: "internal", Message: "server error"}
	if IsNotFound(err500) {
		t.Fatal("expected IsNotFound to be false for 500")
	}

	if IsNotFound(nil) {
		t.Fatal("expected IsNotFound to be false for nil")
	}
}
