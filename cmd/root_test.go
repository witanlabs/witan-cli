package cmd

import (
	"os"
	"path/filepath"
	"testing"
)

func TestResolveStateless_ForcesWithoutCredentials(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
	})

	apiKey = ""
	apiURL = ""
	stateless = false

	t.Setenv("WITAN_API_KEY", "")
	t.Setenv("WITAN_STATELESS", "")
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())

	if !resolveStateless() {
		t.Fatal("expected stateless mode to be forced when no credentials are available")
	}

	key, err := resolveAPIKey()
	if err != nil {
		t.Fatalf("resolveAPIKey returned error: %v", err)
	}
	if key != "" {
		t.Fatalf("expected empty API key in forced stateless mode, got %q", key)
	}
}

func TestResolveStateless_DoesNotForceWithAPIKey(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
	})

	apiKey = ""
	apiURL = ""
	stateless = false

	t.Setenv("WITAN_API_KEY", "test-key")
	t.Setenv("WITAN_STATELESS", "")
	t.Setenv("WITAN_CONFIG_DIR", t.TempDir())

	if resolveStateless() {
		t.Fatal("expected stateful mode when API key is present and stateless is not requested")
	}

	key, err := resolveAPIKey()
	if err != nil {
		t.Fatalf("resolveAPIKey returned error: %v", err)
	}
	if key != "test-key" {
		t.Fatalf("expected API key from environment, got %q", key)
	}
}

func TestResolveStateless_ForcesWhenConfigLoadErrors(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
	})

	apiKey = ""
	apiURL = ""
	stateless = false

	t.Setenv("WITAN_API_KEY", "")
	t.Setenv("WITAN_STATELESS", "")

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)
	if err := os.Mkdir(filepath.Join(configDir, "config.json"), 0o755); err != nil {
		t.Fatalf("setup invalid config path: %v", err)
	}

	if !resolveStateless() {
		t.Fatal("expected stateless mode to be forced when config cannot be loaded")
	}
}

func TestResolveAPIKey_AllowsStatelessFallbackWhenConfigLoadErrors(t *testing.T) {
	origAPIKey := apiKey
	origAPIURL := apiURL
	origStateless := stateless
	t.Cleanup(func() {
		apiKey = origAPIKey
		apiURL = origAPIURL
		stateless = origStateless
	})

	apiKey = ""
	apiURL = ""
	stateless = false

	t.Setenv("WITAN_API_KEY", "")
	t.Setenv("WITAN_STATELESS", "")

	configDir := t.TempDir()
	t.Setenv("WITAN_CONFIG_DIR", configDir)
	if err := os.Mkdir(filepath.Join(configDir, "config.json"), 0o755); err != nil {
		t.Fatalf("setup invalid config path: %v", err)
	}

	key, err := resolveAPIKey()
	if err != nil {
		t.Fatalf("expected stateless fallback, got error: %v", err)
	}
	if key != "" {
		t.Fatalf("expected empty API key in forced stateless mode, got %q", key)
	}
}
