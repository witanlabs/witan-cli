package cmd

import (
	"encoding/json"
	"fmt"
	"net/http"
	"net/url"
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

const versionHealthRequestTimeout = 5 * time.Second

var rootCmd = &cobra.Command{
	Use:   "witan",
	Short: "Witan CLI - spreadsheet tools for agents",
	Long: `Witan CLI provides spreadsheet workflows for calculation, script-driven read/write automation, linting, and rendering.

Workflows:
  auth     Sign in, inspect auth status, or sign out for organization-backed requests.
  read     Extract text from documents (PDF, DOCX, PPTX, HTML, text).
  xlsx     Recalculate formulas, run read/write scripts, lint formulas, and render ranges.

Modes:
  Stateful (default when authenticated):
    Uploads workbook revisions and reuses them across commands.
  Stateless (--stateless, or when no credentials are available):
    Sends the workbook with each request and keeps no server-side file cache.

Quick start:
  witan auth login
  witan auth status
  witan read report.pdf --outline
  witan read report.pdf --pages 1-5
  witan xlsx calc report.xlsx
  witan xlsx exec report.xlsx --expr 'await xlsx.readCell(wb, "Summary!A1")'
  witan xlsx lint report.xlsx --skip-rule D001
  witan xlsx render report.xlsx -r "Sheet1!A1:F20" -o preview.png

Limits:
  Workbook inputs must be 25 MB or smaller.`,
	Version:       Version,
	SilenceErrors: true,
}

func init() {
	cobra.AddTemplateFunc("witanVersionDetails", versionDetails)
	rootCmd.SetVersionTemplate(`{{witanVersionDetails .}}`)

	rootCmd.PersistentFlags().StringVar(&apiKey, "api-key", "", "API key for Witan requests (env: WITAN_API_KEY)")
	rootCmd.PersistentFlags().StringVar(&apiURL, "api-url", "", "Override the Witan API base URL (env: WITAN_API_URL)")
	rootCmd.PersistentFlags().BoolVar(&stateless, "stateless", false, "Send workbook bytes on every request; do not reuse uploaded revisions (env: WITAN_STATELESS)")
}

type healthResponse struct {
	Meta struct {
		Version string `json:"VERSION"`
	} `json:"meta"`
}

func versionDetails(cmd *cobra.Command) string {
	return formatVersionDetails(cmd.DisplayName(), cmd.Version, resolveAPIURL())
}

func formatVersionDetails(displayName, cliVersion, baseURL string) string {
	name := strings.TrimSpace(displayName)
	if name == "" {
		name = "witan"
	}

	version := strings.TrimSpace(cliVersion)
	if version == "" {
		version = "dev"
	}

	var b strings.Builder
	fmt.Fprintf(&b, "%s version %s\n", name, version)

	apiVersion, err := fetchHealthVersion(baseURL)
	if err != nil || apiVersion == "" {
		b.WriteString("API version: unavailable\n")
		return b.String()
	}

	fmt.Fprintf(&b, "API version: %s\n", apiVersion)
	return b.String()
}

func fetchHealthVersion(baseURL string) (string, error) {
	req, err := http.NewRequest("GET", strings.TrimRight(baseURL, "/")+"/health", nil)
	if err != nil {
		return "", err
	}
	setCLIUserAgent(req)

	httpClient := &http.Client{Timeout: versionHealthRequestTimeout}
	resp, err := httpClient.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("HTTP %d", resp.StatusCode)
	}

	var result healthResponse
	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return "", err
	}
	if strings.TrimSpace(result.Meta.Version) == "" {
		return "", fmt.Errorf("missing meta.VERSION in response")
	}
	return result.Meta.Version, nil
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

func resolveRawAPIKey() string {
	if apiKey != "" {
		return apiKey
	}
	return os.Getenv("WITAN_API_KEY")
}

func resolveAuth() (string, string, error) {
	// Priority 1: Raw API key from flag/env
	if rawKey := resolveRawAPIKey(); rawKey != "" {
		orgID, err := resolveAPIKeyOrgID(rawKey)
		if err != nil {
			return "", "", err
		}
		return rawKey, orgID, nil
	}

	// Priority 2: Session token
	cfg, err := config.Load()
	if err != nil {
		if resolveStateless() {
			return "", "", nil
		}
		return "", "", fmt.Errorf("loading auth config: %w", err)
	}
	if cfg.SessionToken == "" {
		if resolveStateless() {
			return "", "", nil
		}
		return "", "", fmt.Errorf("not authenticated: run 'witan auth login' or set --api-key / WITAN_API_KEY")
	}

	jwt, err := exchangeSessionForJWT(resolveManagementAPIURL(), cfg.SessionToken)
	if err != nil {
		return "", "", fmt.Errorf("authentication failed (%v): run 'witan auth login' to re-authenticate", err)
	}
	return jwt, cfg.SessionOrgID, nil
}

// resolveAPIKeyOrgID resolves the org ID for an API key, using the config cache
// or falling back to the management API.
func resolveAPIKeyOrgID(rawAPIKey string) (string, error) {
	cfg, err := config.Load()
	if err == nil {
		if orgID := cfg.OrgIDForAPIKey(rawAPIKey); orgID != "" {
			return orgID, nil
		}
	}

	orgs, err := listOrgsByAPIKey(resolveManagementAPIURL(), rawAPIKey)
	if err != nil {
		return "", fmt.Errorf("resolving org for API key: %w", err)
	}
	if len(orgs) == 0 {
		return "", fmt.Errorf("no organizations found for this API key")
	}

	orgID := orgs[0].ID

	// Best-effort cache in config
	cfg2, loadErr := config.Load()
	if loadErr == nil {
		cfg2.SetOrgIDForAPIKey(rawAPIKey, orgID)
		_ = config.Save(cfg2)
	}

	return orgID, nil
}

// orgEntry represents a single organization from the management API.
type orgEntry struct {
	ID   string `json:"id"`
	Name string `json:"name"`
}

func listOrgsByJWT(mgmtURL, jwt string) ([]orgEntry, error) {
	return listOrgs(mgmtURL, "Bearer "+jwt)
}

func listOrgsByAPIKey(mgmtURL, key string) ([]orgEntry, error) {
	return listOrgs(mgmtURL, "ApiKey "+key)
}

// listOrgs calls GET {mgmtURL}/v0/orgs and returns the list of organizations.
func listOrgs(mgmtURL, authHeader string) ([]orgEntry, error) {
	req, err := http.NewRequest("GET", mgmtURL+"/v0/orgs", nil)
	if err != nil {
		return nil, err
	}
	setCLIUserAgent(req)
	req.Header.Set("Authorization", authHeader)

	httpClient := &http.Client{Timeout: 10 * time.Second}
	resp, err := httpClient.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("HTTP %d", resp.StatusCode)
	}

	var result struct {
		Data []orgEntry `json:"data"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return nil, err
	}
	return result.Data, nil
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
	if derived := deriveManagementAPIURL(resolveAPIURL()); derived != "" {
		return derived
	}
	return "https://management-api.witanlabs.com"
}

func deriveManagementAPIURL(apiBase string) string {
	parsed, err := url.Parse(strings.TrimSpace(apiBase))
	if err != nil || parsed.Scheme == "" || parsed.Host == "" {
		return ""
	}

	host := parsed.Hostname()
	if !strings.HasPrefix(host, "api.") {
		return ""
	}
	rootHost := strings.TrimPrefix(host, "api.")
	if rootHost != "witanlabs.com" && !strings.HasSuffix(rootHost, ".witanlabs.com") {
		return ""
	}

	parsed.Host = "management-api." + rootHost
	if port := parsed.Port(); port != "" {
		parsed.Host += ":" + port
	}
	parsed.Path = ""
	parsed.RawPath = ""
	parsed.RawQuery = ""
	parsed.Fragment = ""
	return strings.TrimRight(parsed.String(), "/")
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

func newAPIClient(bearerToken, orgID string) *client.Client {
	c := client.New(resolveAPIURL(), bearerToken, orgID, resolveStateless())
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
