package cmd

import (
	"net/url"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	sheetsLintRanges   []string
	sheetsLintSkipRule []string
	sheetsLintOnlyRule []string
)

var sheetsLintCmd = &cobra.Command{
	Use:   "lint <spreadsheet>",
	Short: "Run semantic formula checks on a Google Sheet",
	Long: `Run semantic formula checks on a Google Sheets spreadsheet and report diagnostics.

Spreadsheet reference:
  You can reference spreadsheets using either format:
  - Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  - Short form: gs://SPREADSHEET_ID

Behavior:
  - Checks the entire spreadsheet by default.
  - Use one or more --range values to limit which cells are analyzed.
  - Range scopes analysis, not fetching; the API still loads all projection sheets.
  - Runs against the live sheet (no revision parameter).
  - Returns exit code 2 when any Error or Warning is reported.
  - Use --json for machine-readable results.

` + LintRulesHelp + `

Examples:
  witan gsheets lint gs://SPREADSHEET_ID
  witan gsheets lint gs://ID -r "Sheet1!A1:Z50"
  witan gsheets lint gs://ID --skip-rule D003
  witan gsheets lint gs://ID --only-rule D004 --json`,
	Args: cobra.ExactArgs(1),
	RunE: runSheetsLint,
}

func init() {
	sheetsLintCmd.SilenceUsage = true
	sheetsLintCmd.Flags().StringArrayVarP(&sheetsLintRanges, "range", "r", nil, `Sheet-qualified range to lint (repeatable)`)
	sheetsLintCmd.Flags().StringArrayVarP(&sheetsLintSkipRule, "skip-rule", "s", nil, `Rule ID to skip (repeatable)`)
	sheetsLintCmd.Flags().StringArrayVar(&sheetsLintOnlyRule, "only-rule", nil, `Run only these rule IDs (repeatable)`)
	gsheetsCmd.AddCommand(sheetsLintCmd)
}

func runSheetsLint(cmd *cobra.Command, args []string) error {
	spreadsheetRef := args[0]

	if err := validateSheetsRef(spreadsheetRef); err != nil {
		return err
	}

	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}

	params := url.Values{}
	for _, r := range sheetsLintRanges {
		params.Add("range", r)
	}
	for _, r := range sheetsLintSkipRule {
		params.Add("skipRule", r)
	}
	for _, r := range sheetsLintOnlyRule {
		params.Add("onlyRule", r)
	}

	spreadsheetID := client.ExtractSpreadsheetID(spreadsheetRef)

	result, err := auth.Client.GSheetsLint(spreadsheetID, params)
	if err != nil {
		return handleSheetsOpError(err, spreadsheetID, gsheetsJSONOutput)
	}

	return outputLintResult(result, gsheetsJSONOutput)
}
