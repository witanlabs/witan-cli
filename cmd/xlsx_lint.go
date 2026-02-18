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

var lintCmd = &cobra.Command{
	Use:   "lint <file>",
	Short: "Run semantic formula analysis",
	Long: `Run semantic analysis on formulas to catch common bugs that editing tools can't detect.

Examples:
  witan xlsx lint report.xlsx                         # Lint entire workbook
  witan xlsx lint report.xlsx -r "Sheet1!A1:Z50"      # Lint specific range
  witan xlsx lint report.xlsx --skip-rule D031         # Skip spell check`,
	Args: cobra.ExactArgs(1),
	RunE: runLint,
}

func init() {
	lintCmd.Flags().StringArrayVarP(&lintRanges, "range", "r", nil, `Range(s) to lint (repeatable)`)
	lintCmd.Flags().StringArrayVarP(&lintSkipRule, "skip-rule", "s", nil, `Rule IDs to skip (repeatable)`)
	lintCmd.Flags().StringArrayVar(&lintOnlyRule, "only-rule", nil, `Only run these rules (repeatable)`)
	xlsxCmd.AddCommand(lintCmd)
}

func runLint(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

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
