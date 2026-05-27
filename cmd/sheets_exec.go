package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	sheetsExecCode           string
	sheetsExecScript         string
	sheetsExecStdin          bool
	sheetsExecExpr           string
	sheetsExecInputJSON      string
	sheetsExecLocale         string
	sheetsExecTitle          string
	sheetsExecStdinTimeoutMS int
	sheetsExecTimeoutMS      int
	sheetsExecMaxOutputChars int
	sheetsExecCreate         bool
)

const defaultSheetsExecStdinTimeoutMS = 2000

var sheetsExecCmd = &cobra.Command{
	Use:   "exec [<spreadsheet>]",
	Short: "Execute JavaScript against a Google Sheet",
	Long: `Execute JavaScript against a Google Sheets spreadsheet.

Contract:
  - Provide exactly one code source: --code, --script, --stdin, or --expr.
  - --expr wraps input as: return (<expr>);
  - --expr is for single expressions only; use --code for multi-statement scripts.
  - Script code must evaluate to JSON-serializable result values.

Spreadsheet reference:
  - Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  - Short form: gs://SPREADSHEET_ID
  - Create mode: omit the argument or use new / gs://new (with or without --create)

Inputs:
  - --input-json passes any JSON value to the script as input.
  - --locale sets the workbook execution locale explicitly.
  - --title sets the Drive title when creating a new spreadsheet (max 1000 characters).
  - If --input-json is omitted, input defaults to {}.

Defaults:
  - If --locale is omitted, the CLI tries WITAN_LOCALE, then LC_ALL / LC_MESSAGES / LANG.
  - --timeout-ms=0 means no explicit timeout override.
  - --stdin-timeout-ms=2000 aborts --stdin reads that never reach EOF; set 0 to disable.
  - --max-output-chars=0 means no explicit stdout cap override.
  - --create=false means exec expects an existing spreadsheet reference.

Output:
  - Default mode prints stdout first, then:
      - pretty JSON result when ok=true
      - formatted error summary when ok=false
  - --json prints the full response envelope.
    Success shape:
      {"ok":true,"stdout":"...","result":<json>,"writes_detected":<bool>,"accesses":[...],
       "spreadsheet_id":"<id>","url":"https://docs.google.com/..."}
    Failure shape:
      {"ok":false,"stdout":"...","error":{"type":"...","code":"...","message":"..."}}

Behavior:
  - --create starts a new spreadsheet instead of opening an existing one.
  - new and gs://new also select create mode without --create.
  - Changes auto-persist; there is no --save flag.

Exit codes:
  - 0: response has ok=true
  - 1: transport/API error, invalid request, or response has ok=false

Examples:
  witan gsheets exec gs://SPREADSHEET_ID --expr 'await xlsx.readCell(wb, "Sheet1!A1")'
  witan gsheets exec "https://docs.google.com/spreadsheets/d/ID/edit" --script ./script.js
  witan gsheets exec --create --title "Q1 Model" --code 'await xlsx.setCells(wb, [{address: "Sheet1!A1", value: "Hello"}])'
  witan gsheets exec new --stdin <<'WITAN'
  return await xlsx.listSheets(wb);
  WITAN
  cat script.js | witan gsheets exec gs://ID --stdin`,
	Args: validateSheetsExecArgs,
	RunE: runSheetsExec,
}

func init() {
	sheetsExecCmd.SilenceUsage = true
	sheetsExecCmd.Flags().StringVar(&sheetsExecCode, "code", "", "Inline JavaScript source")
	sheetsExecCmd.Flags().StringVar(&sheetsExecScript, "script", "", "Path to a JavaScript file")
	sheetsExecCmd.Flags().BoolVar(&sheetsExecStdin, "stdin", false, "Read JavaScript source from stdin")
	sheetsExecCmd.Flags().StringVar(&sheetsExecExpr, "expr", "", `Single-expression shorthand; wraps as return (<expr>);`)
	sheetsExecCmd.Flags().StringVar(&sheetsExecInputJSON, "input-json", "", "JSON value passed as input to the script")
	sheetsExecCmd.Flags().StringVar(&sheetsExecLocale, "locale", "", "Execution locale (env: WITAN_LOCALE; otherwise LC_ALL / LC_MESSAGES / LANG)")
	sheetsExecCmd.Flags().StringVar(&sheetsExecTitle, "title", "", "Title for a newly created spreadsheet (create mode only, max 1000 characters)")
	sheetsExecCmd.Flags().IntVar(&sheetsExecStdinTimeoutMS, "stdin-timeout-ms", defaultSheetsExecStdinTimeoutMS, "Maximum time to wait for EOF when reading --stdin (0 disables)")
	sheetsExecCmd.Flags().IntVar(&sheetsExecTimeoutMS, "timeout-ms", 0, "Execution timeout in milliseconds (> 0)")
	sheetsExecCmd.Flags().IntVar(&sheetsExecMaxOutputChars, "max-output-chars", 0, "Maximum stdout characters to capture (> 0)")
	sheetsExecCmd.Flags().BoolVar(&sheetsExecCreate, "create", false, "Create a new Google Sheet instead of opening an existing one")
	gsheetsCmd.AddCommand(sheetsExecCmd)
}

func validateSheetsExecArgs(_ *cobra.Command, args []string) error {
	return validateSheetsOpenOrCreateArgs(args, sheetsExecCreate)
}

func runSheetsExec(cmd *cobra.Command, args []string) error {
	create := resolveSheetsCreate(sheetsExecCreate, args)
	if !create && len(args) != 1 {
		return fmt.Errorf("requires exactly 1 spreadsheet reference")
	}

	if err := validateSheetsTitle(sheetsExecTitle, create); err != nil {
		return err
	}
	if sheetsExecTitle != "" && !create {
		return fmt.Errorf("--title can only be used with --create or spreadsheet reference new")
	}

	if err := validateExecPositiveFlag(cmd, "timeout-ms", sheetsExecTimeoutMS); err != nil {
		return err
	}
	if err := validateExecNonNegativeFlag(cmd, "stdin-timeout-ms", sheetsExecStdinTimeoutMS); err != nil {
		return err
	}
	if err := validateExecPositiveFlag(cmd, "max-output-chars", sheetsExecMaxOutputChars); err != nil {
		return err
	}

	code, err := resolveExecCodeSource(cmd, os.Stdin, sheetsExecCode, sheetsExecScript, sheetsExecStdin, sheetsExecExpr, sheetsExecStdinTimeoutMS)
	if err != nil {
		return err
	}
	if strings.TrimSpace(code) == "" {
		return fmt.Errorf("exec code must not be empty")
	}

	input, err := parseExecInput(sheetsExecInputJSON, cmd.Flags().Changed("input-json"))
	if err != nil {
		return err
	}

	locale, err := resolveLocale(cmd, "locale", sheetsExecLocale, true, true)
	if err != nil {
		return err
	}

	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}

	req := client.ExecRequest{
		Code:           code,
		Input:          input,
		Title:          sheetsExecTitle,
		Locale:         locale,
		TimeoutMS:      sheetsExecTimeoutMS,
		MaxOutputChars: sheetsExecMaxOutputChars,
	}

	var result *client.ExecResponse
	var spreadsheetID string
	if create {
		result, err = auth.Client.GSheetsExecCreate(req)
	} else {
		spreadsheetID = client.ExtractSpreadsheetID(args[0])
		result, err = auth.Client.GSheetsExec(spreadsheetID, req)
	}
	if err != nil {
		return handleSheetsOpError(err, spreadsheetID, gsheetsJSONOutput)
	}

	if err := outputExecResult(result, gsheetsJSONOutput, formatSheetsExecError); err != nil {
		return err
	}

	if create && result.Ok && !gsheetsJSONOutput {
		outputSheetsCreateHints(result.SpreadsheetID, result.URL, sheetsExecTitle)
	}
	return nil
}

// formatSheetsExecError formats an ExecError for Google Sheets with special handling
// for google-specific error codes.
func formatSheetsExecError(execErr *client.ExecError) string {
	if execErr == nil {
		return "execution failed"
	}

	switch execErr.Code {
	case "google_auth_required":
		return "Google Sheets requires authorization. Run 'witan gsheets connect' to enable access."
	case "google_sheets_not_found":
		return "spreadsheet not found or not shared with your account"
	case "google_sheets_forbidden":
		return "you don't have permission to access this spreadsheet"
	}

	return formatExecError(execErr)
}
