package cmd

import (
	"context"
	"errors"
	"fmt"
	"net/http"
	"os"
	"os/signal"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
	"github.com/witanlabs/witan-cli/config"
)

var sheetsStatusWait bool

var sheetsStatusCmd = &cobra.Command{
	Use:   "status [<spreadsheet>]",
	Short: "Check Google Sheets connection or per-sheet authorization status",
	Long: `Check Google Sheets status.

Without an argument, reports account connection status. Connecting does not
grant access to existing sheets — use 'witan gsheets authorize' for that.

With a spreadsheet argument, reports whether that specific sheet is authorized.

Spreadsheet reference:
  - Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  - Short form: gs://SPREADSHEET_ID

Flags:
  --wait   Poll until connected (no argument) or authorized (with argument).
           Use this after handing a connect/authorize URL to a human.

Examples:
  witan gsheets status
  witan gsheets status --json
  witan gsheets status gs://SPREADSHEET_ID
  witan gsheets status gs://SPREADSHEET_ID --wait --json`,
	Args: cobra.MaximumNArgs(1),
	RunE: runSheetsStatus,
}

func init() {
	sheetsStatusCmd.SilenceUsage = true
	sheetsStatusCmd.Flags().BoolVar(&sheetsStatusWait, "wait", false, "Poll until connected/authorized")
	gsheetsCmd.AddCommand(sheetsStatusCmd)
}

type sheetsStatusReport struct {
	Status    string `json:"status"`
	ExpiresAt string `json:"expires_at,omitempty"`
	Error     string `json:"error,omitempty"`
	Hint      string `json:"hint,omitempty"`
}

// sheetStatusOutput is the machine-readable per-sheet authorization result.
type sheetStatusOutput struct {
	Authorized bool `json:"authorized"`
}

func runSheetsStatus(cmd *cobra.Command, args []string) error {
	if len(args) == 1 {
		return runSheetStatusForSheet(cmd, args[0])
	}
	return runSheetsConnectionStatus(cmd)
}

func runSheetsConnectionStatus(cmd *cobra.Command) error {
	if sheetsStatusWait {
		return waitForConnection(cmd)
	}

	report := inspectSheetsStatus()

	if gsheetsJSONOutput {
		return jsonPrintTo(cmd.OutOrStdout(), report)
	}

	printSheetsStatus(cmd, report)
	return nil
}

func waitForConnection(cmd *cobra.Command) error {
	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}
	httpClient := &http.Client{Timeout: 30 * time.Second}

	ctx, cancel := signal.NotifyContext(context.Background(), os.Interrupt)
	defer cancel()

	var lastErr error
	err = pollUntil(ctx, connectPollInterval, connectTimeout,
		timeoutMessage("timed out waiting for connection", &lastErr),
		func() (bool, error) {
			status, err := getGoogleSheetsIntegrationStatus(httpClient, auth.MgmtURL, auth.JWT)
			if err != nil {
				if isTransientManagementError(err) {
					lastErr = err
					return false, nil // transient — keep polling
				}
				return false, sheetsStatusCheckError(err)
			}
			return status.isActive(), nil
		})
	if err != nil {
		return err
	}

	// Emit the same shape as `status` without --wait, so agents that poll and
	// check `status == "connected"` (per the skill docs) work on both paths.
	report := inspectSheetsStatus()
	if gsheetsJSONOutput {
		return jsonPrintTo(cmd.OutOrStdout(), report)
	}
	printSheetsStatus(cmd, report)
	return nil
}

func runSheetStatusForSheet(cmd *cobra.Command, ref string) error {
	if isSheetsCreateRef(ref) {
		return fmt.Errorf("'new' is not a real spreadsheet reference")
	}
	if err := validateSheetsRef(ref); err != nil {
		return err
	}

	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}
	httpClient := &http.Client{Timeout: 30 * time.Second}
	spreadsheet := client.ExtractSpreadsheetID(ref)

	check := func() (bool, error) {
		return authorizeSheetStatus(httpClient, auth.MgmtURL, auth.JWT, spreadsheet)
	}

	var authorized bool
	if sheetsStatusWait {
		ctx, cancel := signal.NotifyContext(context.Background(), os.Interrupt)
		defer cancel()
		var lastErr error
		err = pollUntil(ctx, connectPollInterval, connectTimeout,
			timeoutMessage("timed out waiting for authorization", &lastErr),
			func() (bool, error) {
				a, err := check()
				if err != nil {
					if errors.Is(err, errSheetsAuthUnavailable) {
						lastErr = err
						return false, nil // transient — keep polling
					}
					return false, err
				}
				return a, nil
			})
		if err != nil {
			return err
		}
		authorized = true
	} else {
		authorized, err = check()
		if err != nil {
			return err // includes the transient "could not determine; retry" case
		}
	}

	if gsheetsJSONOutput {
		return jsonPrint(sheetStatusOutput{Authorized: authorized})
	}

	out := cmd.OutOrStdout()
	if authorized {
		fmt.Fprintln(out, "Sheet: authorized")
		return nil
	}
	fmt.Fprintln(out, "Sheet: not authorized")
	fmt.Fprintf(os.Stderr, "Hint: run 'witan gsheets authorize %s'\n", ref)
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
