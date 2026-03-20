package config

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"os"
	"path/filepath"
)

const configVersion = 1

type Config struct {
	Version      int               `json:"v,omitempty"`
	SessionToken string            `json:"session_token,omitempty"`
	SessionOrgID string            `json:"session_org_id,omitempty"`
	APIKeyOrgs   map[string]string `json:"api_key_orgs,omitempty"` // sha256(apiKey) -> orgID
}

// HashAPIKey returns the hex-encoded SHA-256 of an API key.
func HashAPIKey(apiKey string) string {
	h := sha256.Sum256([]byte(apiKey))
	return hex.EncodeToString(h[:])
}

// OrgIDForAPIKey looks up the cached org ID for the given API key.
func (c *Config) OrgIDForAPIKey(apiKey string) string {
	if c.APIKeyOrgs == nil {
		return ""
	}
	return c.APIKeyOrgs[HashAPIKey(apiKey)]
}

// SetOrgIDForAPIKey caches an org ID for the given API key.
func (c *Config) SetOrgIDForAPIKey(apiKey, orgID string) {
	if c.APIKeyOrgs == nil {
		c.APIKeyOrgs = make(map[string]string)
	}
	c.APIKeyOrgs[HashAPIKey(apiKey)] = orgID
}

func dir() (string, error) {
	if v := os.Getenv("WITAN_CONFIG_DIR"); v != "" {
		return v, nil
	}
	if v := os.Getenv("XDG_CONFIG_HOME"); v != "" {
		return filepath.Join(v, "witan"), nil
	}
	home, err := os.UserHomeDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(home, ".config", "witan"), nil
}

func filePath() (string, error) {
	d, err := dir()
	if err != nil {
		return "", err
	}
	return filepath.Join(d, "config.json"), nil
}

// Load reads the config file. Returns a zero-value Config if the file does not
// exist or has an outdated version (the stale file is deleted automatically).
func Load() (Config, error) {
	p, err := filePath()
	if err != nil {
		return Config{}, err
	}
	data, err := os.ReadFile(p)
	if err != nil {
		if os.IsNotExist(err) {
			return Config{}, nil
		}
		return Config{}, err
	}
	var cfg Config
	if err := json.Unmarshal(data, &cfg); err != nil {
		return Config{}, err
	}
	if cfg.Version < configVersion {
		_ = os.Remove(p)
		return Config{}, nil
	}
	return cfg, nil
}

// Save writes the config to disk atomically using a temp file + rename.
func Save(cfg Config) error {
	cfg.Version = configVersion
	p, err := filePath()
	if err != nil {
		return err
	}
	d := filepath.Dir(p)
	if err := os.MkdirAll(d, 0700); err != nil {
		return err
	}
	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	data = append(data, '\n')
	tmp := p + ".tmp"
	if err := os.WriteFile(tmp, data, 0600); err != nil {
		return err
	}
	// Remove dest first for Windows compat (os.Rename fails if dest exists on Windows).
	if err := os.Remove(p); err != nil && !os.IsNotExist(err) {
		_ = os.Remove(tmp)
		return err
	}
	if err := os.Rename(tmp, p); err != nil {
		_ = os.Remove(tmp)
		return err
	}
	return nil
}

// Delete removes the config file.
func Delete() error {
	p, err := filePath()
	if err != nil {
		return err
	}
	err = os.Remove(p)
	if err != nil && os.IsNotExist(err) {
		return nil
	}
	return err
}
