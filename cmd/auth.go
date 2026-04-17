package cmd

import "github.com/spf13/cobra"

var authCmd = &cobra.Command{
	Use:   "auth",
	Short: "Sign in and session management",
	Long: `Manage authentication for Witan CLI.

Use login to start browser sign-in and save a local session.
Use status to inspect which credential is active right now.
Use logout to revoke that session and clear local credentials.

Examples:
  witan auth login
  witan auth status
  witan auth logout`,
}

func init() {
	rootCmd.AddCommand(authCmd)
}
