package client

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sync"
)

// CacheEntry maps a file hash to its uploaded file ID and revision.
type CacheEntry struct {
	FileID     string `json:"file_id"`
	RevisionID string `json:"revision_id"`
	Bytes      int64  `json:"bytes"`
	Filename   string `json:"filename"`
}

// cacheData is the on-disk JSON structure.
type cacheData struct {
	Version int                   `json:"v"`
	Files   map[string]CacheEntry `json:"files"`
	Known   map[string]CacheEntry `json:"known"`
}

// FileCache persists hash→fileId mappings on disk.
// If no writable directory is found, it operates in-memory only.
type FileCache struct {
	mu       sync.Mutex
	dir      string // empty string = in-memory only
	data     cacheData
	inMemory map[string]CacheEntry
}

// NewFileCache probes for a writable cache directory using the cascade:
//  1. $TMPDIR/witan/ (or os.TempDir()/witan/)
//  2. .witan/ in cwd
//  3. in-memory only (no persistence)
func NewFileCache() *FileCache {
	fc := &FileCache{
		inMemory: make(map[string]CacheEntry),
	}

	// Tier 1: tmpdir
	tmpdir := os.TempDir()
	if dir := filepath.Join(tmpdir, "witan"); probeWritable(dir) {
		fc.dir = dir
		fc.load()
		return fc
	}

	// Tier 2: cwd/.witan
	if cwd, err := os.Getwd(); err == nil {
		if dir := filepath.Join(cwd, ".witan"); probeWritable(dir) {
			fc.dir = dir
			fc.load()
			return fc
		}
	}

	// Tier 3: in-memory only
	return fc
}

// Get looks up a cache entry by file hash key.
func (fc *FileCache) Get(key string) (CacheEntry, bool) {
	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		e, ok := fc.data.Files[key]
		return e, ok
	}
	e, ok := fc.inMemory[key]
	return e, ok
}

// Put stores a cache entry and persists to disk if possible.
func (fc *FileCache) Put(key string, entry CacheEntry) {
	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		fc.data.Files[key] = entry
		fc.save()
	} else {
		fc.inMemory[key] = entry
	}
}

// GetKnown looks up a cache entry by local file identity.
func (fc *FileCache) GetKnown(filePath, baseURL string) (CacheEntry, bool) {
	key := KnownFileKey(filePath, baseURL)

	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		e, ok := fc.data.Known[key]
		return e, ok
	}
	e, ok := fc.inMemory[key]
	return e, ok
}

// PutKnown stores a cache entry by local file identity.
func (fc *FileCache) PutKnown(filePath, baseURL string, entry CacheEntry) {
	key := KnownFileKey(filePath, baseURL)

	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		fc.data.Known[key] = entry
		fc.save()
	} else {
		fc.inMemory[key] = entry
	}
}

// Evict removes a cache entry (e.g. after a 404).
func (fc *FileCache) Evict(key string) {
	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		delete(fc.data.Files, key)
		fc.save()
	} else {
		delete(fc.inMemory, key)
	}
}

// EvictKnown removes a local-file cache entry.
func (fc *FileCache) EvictKnown(filePath, baseURL string) {
	key := KnownFileKey(filePath, baseURL)

	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		delete(fc.data.Known, key)
		fc.save()
	} else {
		delete(fc.inMemory, key)
	}
}

// HashFile computes the cache key for a local file: "sha256:<hex>@<baseURL>".
func HashFile(filePath, baseURL string) (string, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return "", fmt.Errorf("opening file for hashing: %w", err)
	}
	defer f.Close()

	h := sha256.New()
	if _, err := io.Copy(h, f); err != nil {
		return "", fmt.Errorf("hashing file: %w", err)
	}
	hex := hex.EncodeToString(h.Sum(nil))
	return "sha256:" + hex + "@" + baseURL, nil
}

// KnownFileKey computes the cache key for a local file identity.
func KnownFileKey(filePath, baseURL string) string {
	absPath, err := filepath.Abs(filePath)
	if err != nil {
		absPath = filePath
	}
	return "path:" + filepath.Clean(absPath) + "@" + baseURL
}

func (fc *FileCache) load() {
	path := filepath.Join(fc.dir, "cache.json")
	raw, err := os.ReadFile(path)
	if err != nil {
		fc.data = cacheData{
			Version: 2,
			Files:   make(map[string]CacheEntry),
			Known:   make(map[string]CacheEntry),
		}
		return
	}
	if err := json.Unmarshal(raw, &fc.data); err != nil || fc.data.Version < 1 || fc.data.Version > 2 {
		fc.data = cacheData{
			Version: 2,
			Files:   make(map[string]CacheEntry),
			Known:   make(map[string]CacheEntry),
		}
		return
	}
	if fc.data.Files == nil {
		fc.data.Files = make(map[string]CacheEntry)
	}
	if fc.data.Known == nil {
		fc.data.Known = make(map[string]CacheEntry)
	}
	fc.data.Version = 2
}

func (fc *FileCache) save() {
	if fc.dir == "" {
		return
	}
	_ = os.MkdirAll(fc.dir, 0o755)
	raw, err := json.MarshalIndent(fc.data, "", "  ")
	if err != nil {
		return
	}
	_ = os.WriteFile(filepath.Join(fc.dir, "cache.json"), raw, 0o644)
}

// probeWritable tries to create the directory and write a probe file.
func probeWritable(dir string) bool {
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return false
	}
	probe := filepath.Join(dir, ".probe")
	if err := os.WriteFile(probe, []byte("ok"), 0o644); err != nil {
		return false
	}
	os.Remove(probe)
	return true
}
