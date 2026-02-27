package cmd

import (
	"bytes"
	"fmt"
	"net/http"
	"os"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/config"
)

var logoutCmd = &cobra.Command{
	Use:   "logout",
	Short: "Log out of Witan",
	Long: `Sign out from Witan on this machine.

What happens:
  - Attempts to revoke the current server session (best effort).
  - Removes locally saved session credentials.
  - If no session exists, prints "Not logged in." and exits successfully.

Example:
  witan auth logout`,
	RunE: runLogout,
}

func init() {
	logoutCmd.SilenceUsage = true
	authCmd.AddCommand(logoutCmd)
}

func runLogout(cmd *cobra.Command, args []string) error {
	cfg, err := config.Load()
	if err != nil {
		return fmt.Errorf("failed to load config: %w", err)
	}

	if cfg.SessionToken == "" {
		fmt.Fprintln(os.Stderr, "Not logged in.")
		return nil
	}

	// Revoke session server-side (best effort)
	mgmtURL := resolveManagementAPIURL()
	httpClient := &http.Client{Timeout: 10 * time.Second}
	req, err := http.NewRequest("POST", mgmtURL+"/v0/auth/sign-out", bytes.NewReader(nil))
	if err != nil {
		fmt.Fprintf(os.Stderr, "Warning: could not revoke session: %v\n", err)
	} else {
		setCLIUserAgent(req)
		req.Header.Set("Authorization", "Bearer "+cfg.SessionToken)
		resp, err := httpClient.Do(req)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Warning: could not revoke session: %v\n", err)
		} else {
			resp.Body.Close()
		}
	}

	// Delete local config
	if err := config.Delete(); err != nil {
		return fmt.Errorf("failed to delete config: %w", err)
	}

	fmt.Fprintln(os.Stderr, "\u2713 Logged out")
	return nil
}
