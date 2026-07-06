package cmd

import (
	"fmt"
	"net/url"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	pptxLintSlides   []int
	pptxLintSkipRule []string
	pptxLintOnlyRule []string
)

var pptxLintCmd = &cobra.Command{
	Use:   "lint <file.pptx>",
	Short: "Run semantic presentation checks",
	Long: `Run semantic presentation checks and report diagnostics by severity.

Behavior:
  - Checks the entire presentation by default.
  - Use one or more --slide values to limit analysis.
  - Returns exit code 2 when any Error or Warning is reported.
  - Use --json for machine-readable results.

Pptx-specific rules use the P### namespace; the chart data-integrity family
shared with xlsx lint keeps its D### ids.

Available rules:
  D100 (Error): Chart series data reference fails to resolve
  D101 (Warning): Chart displays cached data that is out of date with its embedded workbook
  D102 (Warning): Chart series data references have mismatched lengths
  D103 (Warning): Chart series renders no data points
  D104 (Warning): Chart series data contains calculation errors
  D105 (Warning): Chart value range contains non-numeric text
  D106 (Warning): Chart data lies outside the explicit axis bounds
  D107 (Error): Chart data contains non-positive values on a logarithmic axis
  D108 (Warning): Pie or doughnut chart plots negative values as positive slices
  D109 (Warning): Scatter or bubble chart has non-numeric X values
  D110 (Warning): Chart has multiple series plotting the same values range

Examples:
  witan pptx lint deck.pptx
  witan pptx lint deck.pptx -p 1 -p 3
  witan pptx lint deck.pptx --skip-rule D101
  witan pptx lint deck.pptx --only-rule D100 --only-rule D101`,
	Args: cobra.ExactArgs(1),
	RunE: runPPTXLint,
}

func init() {
	pptxLintCmd.Flags().IntSliceVarP(&pptxLintSlides, "slide", "p", nil, `1-based slide number to lint (repeatable)`)
	pptxLintCmd.Flags().StringArrayVarP(&pptxLintSkipRule, "skip-rule", "s", nil, `Rule ID to skip (repeatable)`)
	pptxLintCmd.Flags().StringArrayVar(&pptxLintOnlyRule, "only-rule", nil, `Run only these rule IDs (repeatable)`)
	pptxCmd.AddCommand(pptxLintCmd)
}

func runPPTXLint(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

	if strings.ToLower(filepath.Ext(filePath)) != ".pptx" {
		return fmt.Errorf("PPTX path must end in .pptx")
	}

	key, orgID, err := resolveAuth()
	if err != nil {
		return err
	}

	c := newAPIClient(key, orgID)

	// Build query params with repeated values
	params := url.Values{}
	for _, s := range pptxLintSlides {
		params.Add("slide", strconv.Itoa(s))
	}
	for _, r := range pptxLintSkipRule {
		params.Add("skipRule", r)
	}
	for _, r := range pptxLintOnlyRule {
		params.Add("onlyRule", r)
	}

	var result *client.PptxLintResponse
	if c.Stateless {
		result, err = c.PPTXLint(filePath, params)
	} else {
		var fileID, revisionID string
		fileID, revisionID, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesPPTXLint(fileID, revisionID, params)
			if client.IsNotFound(err) {
				fileID, revisionID, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesPPTXLint(fileID, revisionID, params)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	if pptxJSONOutput {
		if err := jsonPrint(result); err != nil {
			return err
		}
	} else {
		// Group diagnostics by severity
		errors := []client.LintDiagnostic{}
		warnings := []client.LintDiagnostic{}
		infos := []client.LintDiagnostic{}

		for _, d := range result.Diagnostics {
			flat := client.LintDiagnostic{
				Severity: d.Severity,
				RuleId:   d.RuleId,
				Message:  d.Message,
				Location: d.Location,
			}
			switch d.Severity {
			case "Error":
				errors = append(errors, flat)
			case "Warning":
				warnings = append(warnings, flat)
			default:
				infos = append(infos, flat)
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
