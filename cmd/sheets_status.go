package cmd

import (
	"fmt"
	"net/http"
	"os"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/config"
)

var sheetsStatusCmd = &cobra.Command{
	Use:   "status",
	Short: "Check Google Sheets connection status",
	Long: `Check whether Google Sheets is connected to your Witan account.

Reports:
  - Connection status (connected/not connected)
  - Token expiration time (if connected)

Examples:
  witan gsheets status
  witan gsheets status --json`,
	RunE: runSheetsStatus,
}

func init() {
	sheetsStatusCmd.SilenceUsage = true
	gsheetsCmd.AddCommand(sheetsStatusCmd)
}

type sheetsStatusReport struct {
	Status    string `json:"status"`
	ExpiresAt string `json:"expires_at,omitempty"`
	Error     string `json:"error,omitempty"`
	Hint      string `json:"hint,omitempty"`
}

func runSheetsStatus(cmd *cobra.Command, args []string) error {
	report := inspectSheetsStatus()

	if gsheetsJSONOutput {
		return jsonPrintTo(cmd.OutOrStdout(), report)
	}

	printSheetsStatus(cmd, report)
	return nil
}

func inspectSheetsStatus() sheetsStatusReport {
	// Check for API key (not supported)
	if resolveRawAPIKey() != "" {
		return sheetsStatusReport{
			Status: "unavailable",
			Error:  "API key authentication does not support Google Sheets",
			Hint:   "run 'witan auth login' and remove --api-key or WITAN_API_KEY",
		}
	}

	cfg, err := config.Load()
	if err != nil {
		return sheetsStatusReport{
			Status: "unavailable",
			Error:  fmt.Sprintf("loading auth config: %v", err),
			Hint:   "run 'witan auth login'",
		}
	}

	if cfg.SessionToken == "" {
		return sheetsStatusReport{
			Status: "unavailable",
			Error:  "not authenticated",
			Hint:   "run 'witan auth login'",
		}
	}

	mgmtURL := resolveManagementAPIURL()

	// Exchange session token for JWT
	jwt, err := exchangeSessionForJWT(mgmtURL, cfg.SessionToken)
	if err != nil {
		return sheetsStatusReport{
			Status: "unavailable",
			Error:  "session expired",
			Hint:   "run 'witan auth login' to re-authenticate",
		}
	}

	httpClient := &http.Client{Timeout: 10 * time.Second}

	integration, err := getGoogleSheetsIntegrationStatus(httpClient, mgmtURL, jwt)
	if err != nil {
		if apiErr, ok := err.(*ManagementAPIError); ok {
			switch apiErr.Code {
			case "unauthorized":
				return sheetsStatusReport{
					Status: "unavailable",
					Error:  "session expired",
					Hint:   "run 'witan auth login' to re-authenticate",
				}
			case "forbidden":
				return sheetsStatusReport{
					Status: "unavailable",
					Error:  "API key authentication does not support Google Sheets",
					Hint:   "run 'witan auth login' and remove --api-key or WITAN_API_KEY",
				}
			}
		}
		return sheetsStatusReport{
			Status: "unknown",
			Error:  err.Error(),
		}
	}

	if integration.needsReauth() {
		return sheetsStatusReport{
			Status:    "expired",
			ExpiresAt: integration.ExpiresAt,
			Error:     "Google authorization has expired or been revoked",
			Hint:      "run 'witan gsheets connect' to reconnect",
		}
	}

	if integration.isActive() {
		return sheetsStatusReport{
			Status:    "connected",
			ExpiresAt: integration.ExpiresAt,
		}
	}

	return sheetsStatusReport{
		Status: "not_connected",
		Hint:   "run 'witan gsheets connect' to enable Google Sheets access",
	}
}

func printSheetsStatus(cmd *cobra.Command, report sheetsStatusReport) {
	out := cmd.OutOrStdout()

	switch report.Status {
	case "connected":
		fmt.Fprintln(out, "Google Sheets: connected")
		if report.ExpiresAt != "" {
			fmt.Fprintf(out, "Token expires: %s\n", report.ExpiresAt)
		}
	case "not_connected":
		fmt.Fprintln(out, "Google Sheets: not connected")
	case "expired":
		fmt.Fprintln(out, "Google Sheets: expired")
		fmt.Fprintln(os.Stderr, "Error:", report.Error)
	case "unavailable":
		fmt.Fprintln(out, "Google Sheets: unavailable")
		fmt.Fprintln(os.Stderr, "Error:", report.Error)
	default:
		fmt.Fprintf(out, "Google Sheets: %s\n", report.Status)
		if report.Error != "" {
			fmt.Fprintln(os.Stderr, "Error:", report.Error)
		}
	}

	if report.Hint != "" {
		fmt.Fprintln(os.Stderr, "Hint:", report.Hint)
	}
}
