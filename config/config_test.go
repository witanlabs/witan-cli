package config

import (
	"os"
	"path/filepath"
	"testing"
)

func TestLoad_ConfigFileIsDirectory(t *testing.T) {
	tmp := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", tmp)

	cfgPath := filepath.Join(tmp, "config.json")
	if err := os.Mkdir(cfgPath, 0o755); err != nil {
		t.Fatalf("setup config dir: %v", err)
	}

	if _, err := Load(); err == nil {
		t.Fatalf("expected read error when config file is a directory")
	} else if os.IsNotExist(err) {
		t.Fatalf("expected non-ENOENT error, got %v", err)
	}
}
