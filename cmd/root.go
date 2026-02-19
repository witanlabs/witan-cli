package cmd

import (
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/config"
)

// Version is set at build time via -ldflags.
var Version = "dev"

var (
	apiKey    string
	apiURL    string
	stateless bool
)

var rootCmd = &cobra.Command{
	Use:           "witan",
	Short:         "Witan CLI — spreadsheet tools for agents",
	Version:       Version,
	SilenceErrors: true,
}

func init() {
	rootCmd.PersistentFlags().StringVar(&apiKey, "api-key", "", "Witan API key (env: WITAN_API_KEY)")
	rootCmd.PersistentFlags().StringVar(&apiURL, "api-url", "", "Witan API URL (env: WITAN_API_URL)")
	rootCmd.PersistentFlags().BoolVar(&stateless, "stateless", false, "Zero data retention: send file with every request (no upload/caching) (env: WITAN_STATELESS)")
}

func resolveStateless() bool {
	if stateless {
		return true
	}
	v := os.Getenv("WITAN_STATELESS")
	if v == "1" || v == "true" {
		return true
	}
	return !hasAuthCredentials()
}

func resolveAPIKey() (string, error) {
	if apiKey != "" {
		return apiKey, nil
	}
	if v := os.Getenv("WITAN_API_KEY"); v != "" {
		return v, nil
	}
	cfg, err := config.Load()
	if err != nil {
		return "", fmt.Errorf("loading auth config: %w", err)
	}
	if cfg.SessionToken == "" {
		if resolveStateless() {
			return "", nil
		}
		return "", fmt.Errorf("not authenticated: run 'witan auth login' or set --api-key / WITAN_API_KEY")
	}
	jwt, err := exchangeSessionForJWT(resolveManagementAPIURL(), cfg.SessionToken)
	if err != nil {
		return "", fmt.Errorf("authentication failed (%v): run 'witan auth login' to re-authenticate", err)
	}
	return jwt, nil
}

func hasAuthCredentials() bool {
	if apiKey != "" || os.Getenv("WITAN_API_KEY") != "" {
		return true
	}
	cfg, err := config.Load()
	if err != nil {
		return true
	}
	return cfg.SessionToken != ""
}

func resolveManagementAPIURL() string {
	if v := os.Getenv("WITAN_MANAGEMENT_API_URL"); v != "" {
		return v
	}
	return "https://management-api.witanlabs.com"
}

func exchangeSessionForJWT(mgmtURL, sessionToken string) (string, error) {
	req, err := http.NewRequest("GET", mgmtURL+"/v0/auth/token", nil)
	if err != nil {
		return "", err
	}
	req.Header.Set("Authorization", "Bearer "+sessionToken)

	client := &http.Client{Timeout: 10 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("HTTP %d", resp.StatusCode)
	}

	var result struct {
		Token string `json:"token"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return "", err
	}
	if result.Token == "" {
		return "", fmt.Errorf("empty token in response")
	}
	return result.Token, nil
}

func resolveAPIURL() string {
	if apiURL != "" {
		return apiURL
	}
	if v := os.Getenv("WITAN_API_URL"); v != "" {
		return v
	}
	return "https://api.witanlabs.com"
}

func Execute() error {
	return rootCmd.Execute()
}
