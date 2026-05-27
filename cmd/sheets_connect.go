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
		fmt.Fprintln(os.Stderr, "Google Sheets is already connected.")
		fmt.Fprintln(os.Stderr, "Run 'witan gsheets disconnect' first if you want to reconnect.")
		return nil
	}

	// Request OAuth URL
	body, _ := json.Marshal(connectRequest{RedirectURL: cliCallbackURL})
	req, err := http.NewRequest("POST", auth.MgmtURL+"/v0/integrations/google-sheets/connect", bytes.NewReader(body))
	if err != nil {
		return fmt.Errorf("failed to create connect request: %w", err)
	}
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Authorization", "Bearer "+auth.JWT)
	setCLIUserAgent(req)

	resp, err := httpClient.Do(req)
	if err != nil {
		return fmt.Errorf("failed to initiate Google Sheets connection: %w", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return parseConnectError(resp)
	}

	var connResp connectResponse
	if err := json.NewDecoder(resp.Body).Decode(&connResp); err != nil {
		return fmt.Errorf("failed to parse connect response: %w", err)
	}

	// Open browser
	fmt.Fprintln(os.Stderr, "Opening browser to authorize Google Sheets access...")
	fmt.Fprintf(os.Stderr, "If the browser doesn't open, visit:\n  %s\n\n", connResp.RedirectURL)

	if err := openBrowser(connResp.RedirectURL); err != nil {
		// Non-fatal: user can manually visit the URL
		fmt.Fprintf(os.Stderr, "Could not open browser automatically.\n")
	}

	// Poll until the integration becomes active
	fmt.Fprintln(os.Stderr, "Waiting for authorization...")

	ctx, cancel := signal.NotifyContext(context.Background(), os.Interrupt)
	defer cancel()

	deadline := time.Now().Add(connectTimeout)

	for {
		select {
		case <-ctx.Done():
			return fmt.Errorf("interrupted")
		case <-time.After(connectPollInterval):
		}

		if time.Now().After(deadline) {
			return fmt.Errorf("timed out waiting for authorization; run 'witan gsheets connect' to try again")
		}

		status, err := getGoogleSheetsIntegrationStatus(httpClient, auth.MgmtURL, auth.JWT)
		if err != nil {
			return sheetsStatusCheckError(err)
		}

		if status.isActive() {
			fmt.Fprintln(os.Stderr, "\u2713 Google Sheets connected successfully")
			fmt.Fprintln(os.Stderr, "You can now use gsheets commands:")
			fmt.Fprintln(os.Stderr, "  witan gsheets exec gs://SHEET_ID --expr '...'")
			return nil
		}
	}
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
