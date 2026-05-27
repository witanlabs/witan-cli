package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

var sheetsCreateTitle string

var sheetsCreateCmd = &cobra.Command{
	Use:   "create",
	Short: "Create a new Google Sheet",
	Long: `Create a new Google Sheet in your Google Drive.

Returns the spreadsheet URL that can be used with other gsheets commands.

Requirements:
  - You must be logged in with a user session (witan auth login)
  - Google Sheets must be connected (witan gsheets connect)

Examples:
  witan gsheets create
  witan gsheets create --title "My Budget 2024"
  witan gsheets create --json`,
	RunE: runSheetsCreate,
}

func init() {
	sheetsCreateCmd.SilenceUsage = true
	sheetsCreateCmd.Flags().StringVar(&sheetsCreateTitle, "title", "", "Title for the new spreadsheet (max 1000 characters)")
	gsheetsCmd.AddCommand(sheetsCreateCmd)
}

type sheetsCreateOutput struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	Title         string `json:"title"`
	URL           string `json:"url"`
	GSURL         string `json:"gs_url"`
}

func runSheetsCreate(cmd *cobra.Command, args []string) error {
	// Validate title length per spec (max 1000 characters)
	if len(sheetsCreateTitle) > 1000 {
		return fmt.Errorf("--title must be at most 1000 characters")
	}

	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}

	result, err := auth.Client.CreateGoogleSheet(sheetsCreateTitle)
	if err != nil {
		return err
	}

	output := sheetsCreateOutput{
		SpreadsheetID: result.SpreadsheetID,
		Title:         result.Title,
		URL:           result.URL,
		GSURL:         "gs://" + result.SpreadsheetID,
	}

	if gsheetsJSONOutput {
		return jsonPrint(output)
	}

	if output.Title != "" {
		fmt.Fprintf(os.Stderr, "Created new Google Sheet: %s\n", output.Title)
	} else {
		fmt.Fprintln(os.Stderr, "Created new Google Sheet:")
	}
	fmt.Println(output.URL)
	fmt.Fprintf(os.Stderr, "\nUse with gsheets commands:\n")
	fmt.Fprintf(os.Stderr, "  witan gsheets exec %s --expr '...'\n", output.GSURL)

	return nil
}
