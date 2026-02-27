package cmd

import (
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
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
	Use:   "witan",
	Short: "Witan CLI - spreadsheet tools for agents",
	Long: `Witan CLI provides spreadsheet workflows for calculation, script-driven read/write automation, linting, and rendering.

Workflows:
  auth     Sign in or out for organization-backed requests.
  read     Extract text from documents (PDF, DOCX, PPTX, HTML, text).
  xlsx     Recalculate formulas, run read/write scripts, lint formulas, and render ranges.

Modes:
  Stateful (default when authenticated):
    Uploads workbook revisions and reuses them across commands.
  Stateless (--stateless, or when no credentials are available):
    Sends the workbook with each request and keeps no server-side file cache.

Quick start:
  witan auth login
  witan read report.pdf --outline
  witan read report.pdf --pages 1-5
  witan xlsx calc report.xlsx
  witan xlsx exec report.xlsx --expr 'wb.sheet("Summary").cell("A1").value'
  witan xlsx lint report.xlsx --skip-rule D031
  witan xlsx render report.xlsx -r "Sheet1!A1:F20" -o preview.png

Limits:
  Workbook inputs must be 25 MB or smaller.`,
	Version:       Version,
	SilenceErrors: true,
}

func init() {
	rootCmd.PersistentFlags().StringVar(&apiKey, "api-key", "", "API key for Witan requests (env: WITAN_API_KEY)")
	rootCmd.PersistentFlags().StringVar(&apiURL, "api-url", "", "Override the Witan API base URL (env: WITAN_API_URL)")
	rootCmd.PersistentFlags().BoolVar(&stateless, "stateless", false, "Send workbook bytes on every request; do not reuse uploaded revisions (env: WITAN_STATELESS)")
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
		if resolveStateless() {
			return "", nil
		}
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
		return false
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
	setCLIUserAgent(req)
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

func newAPIClient(apiKey string) *client.Client {
	c := client.New(resolveAPIURL(), apiKey, resolveStateless())
	c.UserAgent = cliUserAgent()
	return c
}

func cliUserAgent() string {
	v := strings.TrimSpace(Version)
	if v == "" {
		v = "dev"
	}
	return "witan-cli/" + v
}

func setCLIUserAgent(req *http.Request) {
	req.Header.Set("User-Agent", cliUserAgent())
}

func Execute() error {
	return rootCmd.Execute()
}
