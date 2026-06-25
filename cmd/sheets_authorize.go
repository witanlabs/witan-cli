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
)

// authorizePickerExpirySeconds is how long the Google Picker state JWT is
// valid (the server expires it after 10 minutes). Surfaced to agents so they
// can tell the human to open the link promptly.
const authorizePickerExpirySeconds = 600

var sheetsAuthorizeCmd = &cobra.Command{
	Use:   "authorize <spreadsheet>",
	Short: "Authorize Witan to access a specific Google Sheet",
	Long: `Authorize Witan to access a specific Google Sheet.

Under the drive.file scope, connecting your account grants no access on its
own — each sheet you did not create must be authorized individually by picking
it in Google's file picker. The grant is recorded at Google and persists, so a
sheet only needs to be authorized once (until you disconnect).

Spreadsheet reference:
  - Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  - Short form: gs://SPREADSHEET_ID

Behavior:
  - Interactive terminal: opens the Google file picker in your browser and
    waits until you pick the file.
  - Agent mode (--json, or no terminal): prints the picker URL and returns
    without opening a browser. Hand the URL to a human to open, then poll
    'witan gsheets status <spreadsheet> --wait' until authorized.

Note: sheets you create via Witan are authorized automatically.

Examples:
  witan gsheets authorize gs://SPREADSHEET_ID
  witan gsheets authorize --json gs://SPREADSHEET_ID`,
	Args: cobra.ExactArgs(1),
	RunE: runSheetsAuthorize,
}

func init() {
	sheetsAuthorizeCmd.SilenceUsage = true
	gsheetsCmd.AddCommand(sheetsAuthorizeCmd)
}

// authorizeOutput is the machine-readable result emitted in agent mode.
type authorizeOutput struct {
	Authorized       bool   `json:"authorized"`
	PickerURL        string `json:"picker_url,omitempty"`
	ExpiresInSeconds int    `json:"expires_in_seconds,omitempty"`
}

func runSheetsAuthorize(cmd *cobra.Command, args []string) error {
	ref := args[0]
	if isSheetsCreateRef(ref) {
		return fmt.Errorf("'new' is not authorizable; sheets you create are authorized automatically")
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

	// Idempotent short-circuit: skip the picker if already authorized.
	// A transient "unavailable" is treated as not-yet-known: fall through to start.
	authorized, err := authorizeSheetStatus(httpClient, auth.MgmtURL, auth.JWT, spreadsheet)
	if err != nil && !errors.Is(err, errSheetsAuthUnavailable) {
		return err
	}
	if err == nil && authorized {
		if agentMode(gsheetsJSONOutput) {
			return jsonPrint(authorizeOutput{Authorized: true})
		}
		fmt.Fprintln(os.Stderr, "Sheet is already authorized.")
		return nil
	}

	pickerURL, err := authorizeSheetStart(httpClient, auth.MgmtURL, auth.JWT, spreadsheet, cliCallbackURL)
	if err != nil {
		return err
	}

	// Agent mode: hand the picker URL off to a human; never open a browser or block.
	if agentMode(gsheetsJSONOutput) {
		return jsonPrint(authorizeOutput{
			Authorized:       false,
			PickerURL:        pickerURL,
			ExpiresInSeconds: authorizePickerExpirySeconds,
		})
	}

	// Interactive: open the picker and poll until authorized.
	fmt.Fprintln(os.Stderr, "Opening browser to pick the file...")
	fmt.Fprintf(os.Stderr, "If the browser doesn't open, visit:\n  %s\n\n", pickerURL)
	if err := openBrowser(pickerURL); err != nil {
		fmt.Fprintln(os.Stderr, "Could not open browser automatically.")
	}

	fmt.Fprintln(os.Stderr, "Waiting for you to pick the file...")

	ctx, cancel := signal.NotifyContext(context.Background(), os.Interrupt)
	defer cancel()

	var lastErr error
	err = pollUntil(ctx, connectPollInterval, connectTimeout,
		timeoutMessage("timed out waiting for authorization; run 'witan gsheets authorize' to try again", &lastErr),
		func() (bool, error) {
			authorized, err := authorizeSheetStatus(httpClient, auth.MgmtURL, auth.JWT, spreadsheet)
			if err != nil {
				if errors.Is(err, errSheetsAuthUnavailable) {
					lastErr = err
					return false, nil // transient — keep polling
				}
				return false, err
			}
			return authorized, nil
		})
	if err != nil {
		return err
	}

	fmt.Fprintln(os.Stderr, "✓ Sheet authorized")
	return nil
}
