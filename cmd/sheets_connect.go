package cmd

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/signal"
	"time"

	"github.com/spf13/cobra"
)

const (
	// cliCallbackURL is the hosted callback page that displays a success message
	// and tells the user to return to the CLI.
	cliCallbackURL = "https://app.witanlabs.com/integrations/google-sheets/cli-callback"

	// connectPollInterval is how often we check connection status after the user
	// completes browser authorization.
	connectPollInterval = 2 * time.Second

	// connectTimeout is the maximum time to wait for the user to complete
	// browser authorization.
	connectTimeout = 5 * time.Minute
)

var sheetsConnectCmd = &cobra.Command{
	Use:   "connect",
	Short: "Connect your Google account to Witan",
	Long: `Start browser-based Google authorization to enable Google Sheets access.

What happens:
  1. The CLI opens a Google authorization page in your browser.
  2. After you approve access, the CLI confirms the connection.
  3. You can then use gsheets commands with your Google Sheets.

Behavior:
  - Interactive terminal: opens the Google authorization page in your browser
    and waits until you approve.
  - Agent mode (--json, or no terminal): prints the authorization URL and
    returns without opening a browser. Hand the URL to a human to open, then
    poll 'witan gsheets status --wait' until connected.

Connecting is account-level and grants no access to existing sheets on its own
(drive.file scope) — authorize each sheet you did not create with
'witan gsheets authorize gs://SPREADSHEET_ID'.

Requirements:
  - You must be logged in with a user session (witan auth login)
  - API key authentication is not supported for Google Sheets

Example:
  witan gsheets connect`,
	RunE: runSheetsConnect,
}

func init() {
	sheetsConnectCmd.SilenceUsage = true
	gsheetsCmd.AddCommand(sheetsConnectCmd)
}

type connectRequest struct {
	RedirectURL string `json:"redirect_url"`
}

type connectResponse struct {
	RedirectURL string `json:"redirect_url"`
}

// connectOutput is the machine-readable result emitted in agent mode.
type connectOutput struct {
	Connected        bool   `json:"connected"`
	AuthorizationURL string `json:"authorization_url,omitempty"`
}

func runSheetsConnect(cmd *cobra.Command, args []string) error {
	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}

	httpClient := &http.Client{Timeout: 30 * time.Second}

	// Check if already connected
	status, err := getGoogleSheetsIntegrationStatus(httpClient, auth.MgmtURL, auth.JWT)
	if err != nil {
		return sheetsStatusCheckError(err)
	}
	if status.isActive() {
		if agentMode(gsheetsJSONOutput) {
			return jsonPrint(connectOutput{Connected: true})
		}
		fmt.Fprintln(os.Stderr, "Google Sheets is already connected.")
		fmt.Fprintln(os.Stderr, "Run 'witan gsheets disconnect' first if you want to reconnect.")
		return nil
	}

	// Begin the OAuth connect flow and get the Google consent URL.
	authURL, err := requestConnectURL(httpClient, auth.MgmtURL, auth.JWT)
	if err != nil {
		return err
	}

	// Agent mode: hand the URL off to a human; never open a browser or block.
	if agentMode(gsheetsJSONOutput) {
		return jsonPrint(connectOutput{Connected: false, AuthorizationURL: authURL})
	}

	// Interactive: open the browser and poll until connected.
	fmt.Fprintln(os.Stderr, "Opening browser to authorize Google Sheets access...")
	fmt.Fprintf(os.Stderr, "If the browser doesn't open, visit:\n  %s\n\n", authURL)

	if err := openBrowser(authURL); err != nil {
		// Non-fatal: user can manually visit the URL
		fmt.Fprintf(os.Stderr, "Could not open browser automatically.\n")
	}

	fmt.Fprintln(os.Stderr, "Waiting for authorization...")

	ctx, cancel := signal.NotifyContext(context.Background(), os.Interrupt)
	defer cancel()

	var lastErr error
	err = pollUntil(ctx, connectPollInterval, connectTimeout,
		timeoutMessage("timed out waiting for authorization; run 'witan gsheets connect' to try again", &lastErr),
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

	fmt.Fprintln(os.Stderr, "\u2713 Google Sheets connected successfully")
	fmt.Fprintln(os.Stderr, "Note: connecting does not grant access to existing sheets.")
	fmt.Fprintln(os.Stderr, "Authorize each sheet you didn't create with:")
	fmt.Fprintln(os.Stderr, "  witan gsheets authorize gs://SHEET_ID")
	return nil
}

// requestConnectURL asks the management API to begin the OAuth connect flow
// and returns the Google consent URL to send the user to.
func requestConnectURL(httpClient *http.Client, mgmtURL, jwt string) (string, error) {
	body, _ := json.Marshal(connectRequest{RedirectURL: cliCallbackURL})
	req, err := http.NewRequest("POST", mgmtURL+"/v0/integrations/google-sheets/connect", bytes.NewReader(body))
	if err != nil {
		return "", fmt.Errorf("failed to create connect request: %w", err)
	}
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Authorization", "Bearer "+jwt)
	setCLIUserAgent(req)

	resp, err := httpClient.Do(req)
	if err != nil {
		return "", fmt.Errorf("failed to initiate Google Sheets connection: %w", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return "", parseConnectError(resp)
	}

	var connResp connectResponse
	if err := json.NewDecoder(resp.Body).Decode(&connResp); err != nil {
		return "", fmt.Errorf("failed to parse connect response: %w", err)
	}
	if connResp.RedirectURL == "" {
		return "", fmt.Errorf("connect response missing redirect_url")
	}
	return connResp.RedirectURL, nil
}

func parseConnectError(resp *http.Response) error {
	body, _ := io.ReadAll(resp.Body)
	err := parseManagementAPIError(resp.StatusCode, body)

	// Check for specific error codes that need custom messages
	if apiErr, ok := err.(*ManagementAPIError); ok {
		switch apiErr.Code {
		case "forbidden":
			return fmt.Errorf("Google Sheets integration requires user authentication.\nRun 'witan auth login' and try again without --api-key or WITAN_API_KEY.")
		case "unauthorized":
			return fmt.Errorf("session expired: run 'witan auth login' to re-authenticate")
		}
	}

	return err
}
