package cmd

import (
	"encoding/base64"
	"fmt"
	"net/url"
	"os"
	"sort"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	calcRanges    []string
	calcErrorOnly bool
)

var calcCmd = &cobra.Command{
	Use:   "calc <file>",
	Short: "Recalculate formulas and update cached values",
	Long: `Recalculate formulas in a workbook, update the file with correct cached
formula values, and report results.

Without --range, calculation is seeded from all formula cells (full workbook behavior).
With --range, calculation is seeded from formulas in the provided ranges; downstream
dependents are still recalculated.

Examples:
  witan xlsx calc report.xlsx                        # Recalc all, show errors only
  witan xlsx calc report.xlsx -r "Sheet1!B1:B20"     # Seed calc from a range, show touched values
  witan xlsx calc report.xlsx -r "Sheet1!B1:B20" -r "Summary!A1:H10"  # Multiple seed ranges`,
	Args: cobra.ExactArgs(1),
	RunE: runCalc,
}

func init() {
	calcCmd.Flags().StringArrayVarP(&calcRanges, "range", "r", nil, `Range(s) to seed calculation from and show touched values for (repeatable)`)
	calcCmd.Flags().BoolVar(&calcErrorOnly, "errors-only", false, "Only show errors, skip successful computed values")
	xlsxCmd.AddCommand(calcCmd)
}

func runCalc(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

	key, err := resolveAPIKey()
	if err != nil {
		return err
	}

	c := client.New(resolveAPIURL(), key, resolveStateless())

	// Build query params with repeated address values
	params := url.Values{}
	for _, r := range calcRanges {
		params.Add("address", r)
	}

	var result *client.CalcResponse
	var fileId string
	if c.Stateless {
		result, err = c.Calc(filePath, params)
	} else {
		var revisionId string
		fileId, revisionId, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesCalc(fileId, revisionId, params)
			if client.IsNotFound(err) {
				fileId, revisionId, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesCalc(fileId, revisionId, params)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	// Write back the updated file
	if c.Stateless && result.File != nil {
		// Stateless: file returned inline as base64
		decoded, err := base64.StdEncoding.DecodeString(*result.File)
		if err != nil {
			return fmt.Errorf("decoding updated file: %w", err)
		}
		if err := os.WriteFile(filePath, decoded, 0o644); err != nil {
			return fmt.Errorf("writing updated file: %w", err)
		}
	} else if !c.Stateless && result.RevisionID != nil {
		// Files-backed: download the new revision
		fileBytes, err := c.DownloadFileContent(fileId, *result.RevisionID)
		if err != nil {
			return fmt.Errorf("downloading updated file: %w", err)
		}
		if err := os.WriteFile(filePath, fileBytes, 0o644); err != nil {
			return fmt.Errorf("writing updated file: %w", err)
		}
		if err := c.UpdateCachedRevision(filePath, fileId, *result.RevisionID); err != nil {
			return fmt.Errorf("updating local cache: %w", err)
		}
	}

	if jsonOutput {
		// Nil out File field — it's a huge base64 blob irrelevant to automation
		result.File = nil
		if err := jsonPrint(result); err != nil {
			return err
		}
	} else {
		// Print results
		touchedCount := len(result.Touched)
		errorCount := len(result.Errors)

		if len(calcRanges) > 0 && !calcErrorOnly {
			// Sort touched cells for stable output
			addresses := make([]string, 0, len(result.Touched))
			for addr := range result.Touched {
				addresses = append(addresses, addr)
			}
			sort.Strings(addresses)

			for _, addr := range addresses {
				cell := result.Touched[addr]
				formula := ""
				if cell.Formula != nil {
					formula = *cell.Formula
				}
				// Check if this cell is an error
				isError := false
				for _, e := range result.Errors {
					if e.Address == addr {
						isError = true
						detail := ""
						if e.Detail != nil {
							detail = " ← " + *e.Detail
						}
						fmt.Printf("%-20s %-30s %s%s\n", addr, formula, e.Code, detail)
						break
					}
				}
				if !isError {
					fmt.Printf("%-20s %-30s %s\n", addr, formula, cell.Value)
				}
			}

			fmt.Printf("\n%d cells recalculated", touchedCount)
			if errorCount > 0 {
				fmt.Printf(", %d error", errorCount)
				if errorCount != 1 {
					fmt.Print("s")
				}
			}
			fmt.Println()
		} else {
			// Errors-only output
			if errorCount == 0 {
				fmt.Printf("%d cells recalculated, 0 errors\n", touchedCount)
			} else {
				fmt.Printf("%d error", errorCount)
				if errorCount != 1 {
					fmt.Print("s")
				}
				fmt.Println(":")
				for _, e := range result.Errors {
					formula := ""
					if e.Formula != nil {
						formula = *e.Formula
					}
					detail := ""
					if e.Detail != nil {
						detail = " ← " + *e.Detail
					}
					fmt.Printf("  %-20s %s  %s%s\n", e.Address, formula, e.Code, detail)
				}
			}
		}
	}

	if len(result.Errors) > 0 {
		return &ExitError{Code: 2}
	}
	return nil
}
