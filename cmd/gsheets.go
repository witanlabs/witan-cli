package cmd

import "github.com/spf13/cobra"

var gsheetsJSONOutput bool

var gsheetsCmd = &cobra.Command{
	Use:   "gsheets",
	Short: "Google Sheets integration",
	Long: `Work with Google Sheets spreadsheets.

Commands:
  connect     Connect your Google account to Witan
  authorize   Authorize Witan to access a specific spreadsheet
  disconnect  Remove Google Sheets connection
  status      Check connection status, or per-sheet authorization status
  create      Create a new Google Sheet
  exec        Execute JavaScript against a Google Sheet
  lint        Run semantic formula checks on a Google Sheet
  render      Render a sheet range as an image
  rpc         Run newline-delimited RPC over stdio (supports --create)

Requirements:
  - You must be logged in with a user session (witan auth login)
  - API key authentication is not supported for Google Sheets

Authorization model:
  - 'connect' links your Google account but grants access to NO existing sheets.
  - Each spreadsheet you did not create must be authorized once with
    'witan gsheets authorize <spreadsheet>' (you pick it in Google's file
    picker). The grant persists until you disconnect.
  - Sheets you create via Witan are authorized automatically.
  - Operations on an un-authorized sheet fail with code needs_file_authorization
    and exit code 3.

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
  witan gsheets authorize gs://SPREADSHEET_ID
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
