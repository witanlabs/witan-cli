package cmd

import (
	"fmt"
	"net/url"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	lintRanges   []string
	lintSkipRule []string
	lintOnlyRule []string
)

const lintRulesHelp = `Available rules:
  D001 (Warning): Double counting: same cells contribute multiple times due to overlapping ranges
  D002 (Warning): MATCH/VLOOKUP/HLOOKUP/XLOOKUP with approximate match requires sorted lookup range
  D003 (Warning): Empty cell references may be coerced to 0 or FALSE in numeric/boolean contexts
  D005 (Warning): Numeric aggregate functions ignore text and boolean values
  D006 (Warning): Unintended scalar broadcast in elementwise operations
  D007 (Warning): MATCH/VLOOKUP/HLOOKUP/XLOOKUP with duplicate keys in lookup array returns first match
  D009 (Warning): Mixed percent and non-percent in addition/subtraction
  D030 (Warning): Formula references a non-anchor cell in a merged range
  D031 (Info): Checks spelling of text values in cells`

var lintCmd = &cobra.Command{
	Use:   "lint <file>",
	Short: "Run semantic formula checks",
	Long: `Run semantic formula checks and report diagnostics by severity.

Behavior:
  - Checks the entire workbook by default.
  - Use one or more --range values to limit analysis.
  - Returns exit code 2 when any Error or Warning is reported.
  - Use --json for machine-readable results.

` + lintRulesHelp + `

Examples:
  witan xlsx lint report.xlsx
  witan xlsx lint report.xlsx -r "Sheet1!A1:Z50"
  witan xlsx lint report.xlsx --skip-rule D031
  witan xlsx lint report.xlsx --only-rule D001 --only-rule D030`,
	Args: cobra.ExactArgs(1),
	RunE: runLint,
}

func init() {
	lintCmd.Flags().StringArrayVarP(&lintRanges, "range", "r", nil, `Sheet-qualified range to lint (repeatable)`)
	lintCmd.Flags().StringArrayVarP(&lintSkipRule, "skip-rule", "s", nil, `Rule ID to skip (repeatable)`)
	lintCmd.Flags().StringArrayVar(&lintOnlyRule, "only-rule", nil, `Run only these rule IDs (repeatable)`)
	xlsxCmd.AddCommand(lintCmd)
}

func runLint(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

	filePath, err := fixExcelExtension(filePath)
	if err != nil {
		return err
	}

	key, err := resolveAPIKey()
	if err != nil {
		return err
	}

	c := client.New(resolveAPIURL(), key, resolveStateless())

	// Build query params with repeated values
	params := url.Values{}
	for _, r := range lintRanges {
		params.Add("range", r)
	}
	for _, r := range lintSkipRule {
		params.Add("skipRule", r)
	}
	for _, r := range lintOnlyRule {
		params.Add("onlyRule", r)
	}

	var result *client.LintResponse
	if c.Stateless {
		result, err = c.Lint(filePath, params)
	} else {
		var fileId, revisionId string
		fileId, revisionId, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesLint(fileId, revisionId, params)
			if client.IsNotFound(err) {
				fileId, revisionId, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesLint(fileId, revisionId, params)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	if jsonOutput {
		if err := jsonPrint(result); err != nil {
			return err
		}
	} else {
		// Group diagnostics by severity
		errors := []client.LintDiagnostic{}
		warnings := []client.LintDiagnostic{}
		infos := []client.LintDiagnostic{}

		for _, d := range result.Diagnostics {
			switch d.Severity {
			case "Error":
				errors = append(errors, d)
			case "Warning":
				warnings = append(warnings, d)
			default:
				infos = append(infos, d)
			}
		}

		// Print diagnostics grouped by severity
		printDiagnosticGroup("Error", errors)
		printDiagnosticGroup("Warning", warnings)
		printDiagnosticGroup("Info", infos)

		// Print summary
		fmt.Printf("%d issue", result.Total)
		if result.Total != 1 {
			fmt.Print("s")
		}
		fmt.Printf(" (%d error", len(errors))
		if len(errors) != 1 {
			fmt.Print("s")
		}
		fmt.Printf(", %d warning", len(warnings))
		if len(warnings) != 1 {
			fmt.Print("s")
		}
		fmt.Printf(", %d info)\n", len(infos))
	}

	// Exit 2 when any error- or warning-severity diagnostics exist
	for _, d := range result.Diagnostics {
		if d.Severity == "Error" || d.Severity == "Warning" {
			return &ExitError{Code: 2}
		}
	}
	return nil
}

func printDiagnosticGroup(severity string, diagnostics []client.LintDiagnostic) {
	if len(diagnostics) == 0 {
		return
	}

	fmt.Printf("%s (%d):\n", severity, len(diagnostics))
	for _, d := range diagnostics {
		location := ""
		if d.Location != nil {
			location = *d.Location
		}
		fmt.Printf("  %-6s %-20s %s\n", d.RuleId, location, d.Message)
	}
	fmt.Println()
}
