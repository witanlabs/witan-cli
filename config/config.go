package config

import (
	"encoding/json"
	"os"
	"path/filepath"
)

type Config struct {
	SessionToken string `json:"session_token,omitempty"`
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

// Load reads the config file. Returns a zero-value Config if the file does not exist.
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
	return cfg, nil
}

// Save writes the config to disk atomically using a temp file + rename.
func Save(cfg Config) error {
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
