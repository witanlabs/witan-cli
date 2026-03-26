package config

import (
	"os"
	"path/filepath"
	"testing"
)

func TestLoad_DiscardsOldVersion(t *testing.T) {
	tmp := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", tmp)

	// Write a v0 config (no version field, like old CLI)
	cfgPath := filepath.Join(tmp, "config.json")
	if err := os.WriteFile(cfgPath, []byte(`{"session_token":"old-tok"}`), 0o600); err != nil {
		t.Fatalf("writing old config: %v", err)
	}

	cfg, err := Load()
	if err != nil {
		t.Fatalf("Load failed: %v", err)
	}
	if cfg.SessionToken != "" {
		t.Fatalf("expected empty config after version discard, got token %q", cfg.SessionToken)
	}

	// File should have been deleted
	if _, err := os.Stat(cfgPath); !os.IsNotExist(err) {
		t.Fatalf("expected config file to be deleted, got %v", err)
	}
}

func TestSave_StampsVersion(t *testing.T) {
	tmp := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", tmp)

	if err := Save(Config{SessionToken: "tok", SessionOrgID: "org_1"}); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	cfg, err := Load()
	if err != nil {
		t.Fatalf("Load failed: %v", err)
	}
	if cfg.Version != configVersion {
		t.Fatalf("expected version %d, got %d", configVersion, cfg.Version)
	}
	if cfg.SessionToken != "tok" || cfg.SessionOrgID != "org_1" {
		t.Fatalf("unexpected config: %+v", cfg)
	}
}

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
