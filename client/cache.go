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

const cacheVersion = 3

// CacheEntry records the server-side identity for a local file path,
// plus the content hash at the time the entry was last updated.
type CacheEntry struct {
	FileID      string `json:"file_id"`
	RevisionID  string `json:"revision_id"`
	ContentHash string `json:"content_hash"`
	Bytes       int64  `json:"bytes"`
	Filename    string `json:"filename"`
}

// cacheData is the on-disk JSON structure.
type cacheData struct {
	Version int                   `json:"v"`
	Entries map[string]CacheEntry `json:"entries"`
}

// FileCache persists path→(fileID, revision, contentHash) mappings on disk.
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

	tmpdir := os.TempDir()
	if dir := filepath.Join(tmpdir, "witan"); probeWritable(dir) {
		fc.dir = dir
		fc.load()
		return fc
	}

	if cwd, err := os.Getwd(); err == nil {
		if dir := filepath.Join(cwd, ".witan"); probeWritable(dir) {
			fc.dir = dir
			fc.load()
			return fc
		}
	}

	return fc
}

// Get looks up a cache entry by local file identity.
func (fc *FileCache) Get(filePath, baseURL, orgID string) (CacheEntry, bool) {
	key := entryKey(filePath, baseURL, orgID)

	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		e, ok := fc.data.Entries[key]
		return e, ok
	}
	e, ok := fc.inMemory[key]
	return e, ok
}

// Put stores a cache entry by local file identity.
func (fc *FileCache) Put(filePath, baseURL, orgID string, entry CacheEntry) {
	key := entryKey(filePath, baseURL, orgID)

	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		fc.data.Entries[key] = entry
		fc.save()
	} else {
		fc.inMemory[key] = entry
	}
}

// Evict removes a cache entry by local file identity.
func (fc *FileCache) Evict(filePath, baseURL, orgID string) {
	key := entryKey(filePath, baseURL, orgID)

	fc.mu.Lock()
	defer fc.mu.Unlock()

	if fc.dir != "" {
		delete(fc.data.Entries, key)
		fc.save()
	} else {
		delete(fc.inMemory, key)
	}
}

// hashFile returns "sha256:<hex>" for the file's content.
func hashFile(filePath string) (string, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return "", fmt.Errorf("opening file for hashing: %w", err)
	}
	defer f.Close()

	h := sha256.New()
	if _, err := io.Copy(h, f); err != nil {
		return "", fmt.Errorf("hashing file: %w", err)
	}
	return "sha256:" + hex.EncodeToString(h.Sum(nil)), nil
}

// entryKey returns the cache key for a local file identity.
// Includes path so that distinct files with identical bytes do not collapse
// into one server-side fileID.
func entryKey(filePath, baseURL, orgID string) string {
	absPath, err := filepath.Abs(filePath)
	if err != nil {
		absPath = filePath
	}
	return filepath.Clean(absPath) + "@" + baseURL + "@" + orgID
}

func (fc *FileCache) load() {
	path := filepath.Join(fc.dir, "cache.json")
	raw, err := os.ReadFile(path)
	if err != nil {
		fc.resetData()
		return
	}
	if err := json.Unmarshal(raw, &fc.data); err != nil || fc.data.Version != cacheVersion {
		fc.resetData()
		return
	}
	if fc.data.Entries == nil {
		fc.data.Entries = make(map[string]CacheEntry)
	}
}

func (fc *FileCache) resetData() {
	fc.data = cacheData{
		Version: cacheVersion,
		Entries: make(map[string]CacheEntry),
	}
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
