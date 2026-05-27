package cmd

import (
	"fmt"
	"io"
	"net/http"
	"os"
	"time"

	"github.com/spf13/cobra"
)

var sheetsDisconnectCmd = &cobra.Command{
	Use:   "disconnect",
	Short: "Remove Google Sheets connection",
	Long: `Disconnect Google Sheets from your Witan account.

What happens:
  - Revokes Google authorization
  - Removes stored Google credentials from Witan
  - Google Sheets URLs will no longer work in gsheets commands

Example:
  witan gsheets disconnect`,
	Args: cobra.NoArgs,
	RunE: runSheetsDisconnect,
}

func init() {
	sheetsDisconnectCmd.SilenceUsage = true
	gsheetsCmd.AddCommand(sheetsDisconnectCmd)
}

func runSheetsDisconnect(cmd *cobra.Command, args []string) error {
	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}

	httpClient := &http.Client{Timeout: 30 * time.Second}

	req, err := http.NewRequest("DELETE", auth.MgmtURL+"/v0/integrations/google-sheets", nil)
	if err != nil {
		return fmt.Errorf("failed to create disconnect request: %w", err)
	}
	req.Header.Set("Authorization", "Bearer "+auth.JWT)
	setCLIUserAgent(req)

	resp, err := httpClient.Do(req)
	if err != nil {
		return fmt.Errorf("failed to disconnect Google Sheets: %w", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode == http.StatusOK {
		fmt.Fprintln(os.Stderr, "\u2713 Google Sheets disconnected")
		return nil
	}

	body, _ := io.ReadAll(resp.Body)
	err = parseManagementAPIError(resp.StatusCode, body)

	// Check for specific error codes that need custom handling
	if apiErr, ok := err.(*ManagementAPIError); ok {
		switch apiErr.Code {
		case "google_sheets_not_connected":
			fmt.Fprintln(os.Stderr, "Google Sheets is not connected.")
			return nil
		case "unauthorized":
			return fmt.Errorf("session expired: run 'witan auth login' to re-authenticate")
		case "forbidden":
			return fmt.Errorf("Google Sheets integration requires user authentication.\nRemove --api-key or WITAN_API_KEY and try again.")
		}
	}

	return err
}
