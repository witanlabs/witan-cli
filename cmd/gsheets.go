package cmd

import "github.com/spf13/cobra"

var gsheetsJSONOutput bool

var gsheetsCmd = &cobra.Command{
	Use:   "gsheets",
	Short: "Google Sheets integration",
	Long: `Work with Google Sheets spreadsheets.

Commands:
  connect     Connect your Google account to Witan
  disconnect  Remove Google Sheets connection
  status      Check Google Sheets connection status
  create      Create a new Google Sheet
  exec        Execute JavaScript against a Google Sheet
  lint        Run semantic formula checks on a Google Sheet
  render      Render a sheet range as an image
  rpc         Run newline-delimited RPC over stdio (supports --create)

Requirements:
  - You must be logged in with a user session (witan auth login)
  - API key authentication is not supported for Google Sheets

Spreadsheet references:
  You can reference spreadsheets using either format:
  - Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  - Short form: gs://SPREADSHEET_ID

Note: Changes to Google Sheets auto-persist (no explicit save step needed).

Output:
  default  Human-friendly summaries
  --json   Raw JSON responses for automation

Examples:
  witan gsheets connect
  witan gsheets create --title "My Budget"
  witan gsheets exec gs://SPREADSHEET_ID --expr 'await xlsx.readCell(wb, "Sheet1!A1")'
  witan gsheets exec --create --title "My Budget" --code 'return await xlsx.listSheets(wb)'
  witan gsheets rpc gs://SPREADSHEET_ID
  witan gsheets rpc --create --title "My Budget"
  witan gsheets --json lint "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit"`,
}

func init() {
	gsheetsCmd.PersistentFlags().BoolVar(&gsheetsJSONOutput, "json", false, "Output raw JSON instead of human-formatted summaries")
	rootCmd.AddCommand(gsheetsCmd)
}
