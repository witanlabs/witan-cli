package cmd

import (
	"fmt"
	"sort"

	"github.com/witanlabs/witan-cli/client"
)

// LintRulesHelp is the shared help text describing available lint rules.
const LintRulesHelp = `Available rules:
  D001 (Warning): Double counting: same cells contribute multiple times due to overlapping ranges
  D002 (Warning): MATCH/VLOOKUP/HLOOKUP/XLOOKUP with approximate match requires sorted lookup range
  D003 (Warning): Empty cell references may be coerced to 0 or FALSE in numeric/boolean contexts
  D005 (Warning): Numeric aggregate functions ignore text and boolean values
  D006 (Warning): Unintended scalar broadcast in elementwise operations
  D007 (Warning): MATCH/VLOOKUP/HLOOKUP/XLOOKUP with duplicate keys in lookup array returns first match
  D008 (Error): Mixed currencies in additive/aggregate contexts
  D009 (Warning): Mixed percent and non-percent in addition/subtraction
  D023 (Warning): Currency values mixed with non-currency semantic formats (percent/date/time/text)
  D030 (Warning): Formula references a non-anchor cell in a merged range`

// outputLintResult outputs lint diagnostics in either JSON or human-readable format.
// Returns exit code 2 if any errors or warnings are found.
func outputLintResult(result *client.LintResponse, useJSON bool) error {
	// Group diagnostics by severity
	var errors, warnings, infos []client.LintDiagnostic
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

	if useJSON {
		if err := jsonPrint(result); err != nil {
			return err
		}
	} else {
		// Sort each group by location
		sortDiagnostics := func(diags []client.LintDiagnostic) {
			sort.Slice(diags, func(i, j int) bool {
				locI := ""
				locJ := ""
				if diags[i].Location != nil {
					locI = *diags[i].Location
				}
				if diags[j].Location != nil {
					locJ = *diags[j].Location
				}
				return locI < locJ
			})
		}
		sortDiagnostics(errors)
		sortDiagnostics(warnings)
		sortDiagnostics(infos)

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

	// Exit with code 2 if any errors or warnings
	if len(errors) > 0 || len(warnings) > 0 {
		return &ExitError{Code: 2}
	}
	return nil
}

// printDiagnosticGroup prints a group of diagnostics with the same severity.
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
