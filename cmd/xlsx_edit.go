package cmd

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"os"
	"sort"
	"strconv"
	"strings"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	editFormat string
	editCells  string
)

var editCmd = &cobra.Command{
	Use:   "edit <file> [address=value ...] [flags]",
	Short: "Edit cell values and formulas in a workbook",
	Long: `Set cell values or formulas in a workbook and save the result.

Each edit is specified as address=value. Use a leading = for formulas (double =).

Use --format/-f to apply an Excel number format to all cells. With --format,
bare addresses (no =value) perform format-only edits.

Use --cells to pass a JSON array of edits for full per-cell control (including
per-cell formats). --cells and --format are mutually exclusive, and positional
edit args are not allowed with --cells.

Examples:
  witan xlsx edit report.xlsx "Sheet1!A1=42"
  witan xlsx edit report.xlsx "Sheet1!A1=42" "Sheet1!B2=hello"
  witan xlsx edit report.xlsx "Sheet1!A1==SUM(B1:B10)"   # formula (double =)
  witan xlsx edit report.xlsx "Sheet1!C3=true"            # boolean
  witan xlsx edit report.xlsx "Sheet1!D4=null"            # clear cell
  witan xlsx edit report.xlsx "Sheet1!A1=42" -f "#,##0.00"        # value + format
  witan xlsx edit report.xlsx "Sheet1!A1" -f "0.00%"              # format-only
  witan xlsx edit report.xlsx --cells '[{"address":"Sheet1!A1","value":42,"format":"#,##0.00"}]'`,
	Args: cobra.MinimumNArgs(1),
	RunE: runEdit,
}

func init() {
	editCmd.Flags().StringVarP(&editFormat, "format", "f", "", "Excel number format to apply to all cells")
	editCmd.Flags().StringVar(&editCells, "cells", "", "JSON array of cell edits (full per-cell control)")
	xlsxCmd.AddCommand(editCmd)
}

// parseEditCell parses "Sheet1!A1=42" into an EditCell.
// If the value starts with "=", it's treated as a formula.
// Otherwise: number → bool → null → string.
// When globalFormat is set, bare addresses (no =value) are allowed for format-only edits.
func parseEditCell(arg, globalFormat string) (client.EditCell, error) {
	// Split on the first '=' after '!' so sheet names containing '=' are preserved.
	start := strings.IndexByte(arg, '!')
	if start < 0 {
		start = 0
	}
	idx := strings.IndexByte(arg[start:], '=')
	if idx < 0 {
		// No '=' found — bare address for format-only edit
		if globalFormat == "" {
			return client.EditCell{}, fmt.Errorf("invalid edit %q: expected address=value (use --format for format-only edits)", arg)
		}
		if arg == "" {
			return client.EditCell{}, fmt.Errorf("invalid edit %q: empty address", arg)
		}
		return client.EditCell{Address: arg, Format: globalFormat}, nil
	}
	idx += start
	address := arg[:idx]
	remainder := arg[idx+1:]

	if address == "" {
		return client.EditCell{}, fmt.Errorf("invalid edit %q: empty address", arg)
	}

	// Formula: remainder starts with "=" → strip leading "=" and use formula field
	if strings.HasPrefix(remainder, "=") {
		return client.EditCell{
			Address: address,
			Formula: remainder, // keep the leading "=" as the formula
			Format:  globalFormat,
		}, nil
	}

	// Try number
	if _, err := strconv.ParseFloat(remainder, 64); err == nil {
		return client.EditCell{Address: address, Value: json.RawMessage(remainder), Format: globalFormat}, nil
	}

	// Try boolean
	lower := strings.ToLower(remainder)
	if lower == "true" || lower == "false" {
		return client.EditCell{Address: address, Value: json.RawMessage(lower), Format: globalFormat}, nil
	}

	// Null (clear cell)
	if lower == "null" {
		return client.EditCell{Address: address, Value: json.RawMessage("null"), Format: globalFormat}, nil
	}

	// String — must be JSON-encoded
	raw, _ := json.Marshal(remainder)
	return client.EditCell{Address: address, Value: json.RawMessage(raw), Format: globalFormat}, nil
}

func runEdit(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

	// Parse cell edits
	var cells []client.EditCell
	if editCells != "" {
		if editFormat != "" {
			return fmt.Errorf("--cells and --format are mutually exclusive")
		}
		if len(args) > 1 {
			return fmt.Errorf("positional edit args are not allowed with --cells")
		}
		if err := json.Unmarshal([]byte(editCells), &cells); err != nil {
			return fmt.Errorf("invalid --cells JSON: %w", err)
		}
		if len(cells) == 0 {
			return fmt.Errorf("--cells array must not be empty")
		}
	} else {
		if len(args) < 2 {
			return fmt.Errorf("at least one edit argument is required")
		}
		cells = make([]client.EditCell, 0, len(args)-1)
		for _, arg := range args[1:] {
			cell, err := parseEditCell(arg, editFormat)
			if err != nil {
				return err
			}
			cells = append(cells, cell)
		}
	}

	key, err := resolveAPIKey()
	if err != nil {
		return err
	}

	c := client.New(resolveAPIURL(), key, resolveStateless())

	var result *client.EditResponse
	var fileId string
	if c.Stateless {
		result, err = c.Edit(filePath, cells)
	} else {
		var revisionId string
		fileId, revisionId, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesEdit(fileId, revisionId, cells)
			if client.IsNotFound(err) {
				fileId, revisionId, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesEdit(fileId, revisionId, cells)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	// Write back the updated file
	if c.Stateless && result.File != nil {
		decoded, err := base64.StdEncoding.DecodeString(*result.File)
		if err != nil {
			return fmt.Errorf("decoding updated file: %w", err)
		}
		if err := os.WriteFile(filePath, decoded, 0o644); err != nil {
			return fmt.Errorf("writing updated file: %w", err)
		}
	} else if !c.Stateless && result.RevisionID != nil {
		fileBytes, err := c.DownloadFileContent(fileId, *result.RevisionID)
		if err != nil {
			return fmt.Errorf("downloading updated file: %w", err)
		}
		if err := os.WriteFile(filePath, fileBytes, 0o644); err != nil {
			return fmt.Errorf("writing updated file: %w", err)
		}
	}

	if jsonOutput {
		result.File = nil
		if err := jsonPrint(result); err != nil {
			return err
		}
	} else {
		touchedCount := len(result.Touched)
		errorCount := len(result.Errors)

		if errorCount == 0 {
			fmt.Printf("Edit applied. %d cells recalculated, 0 errors.\n", touchedCount)
		} else {
			fmt.Printf("%d error", errorCount)
			if errorCount != 1 {
				fmt.Print("s")
			}
			fmt.Println(":")

			// Sort errors by address for stable output
			sorted := make([]client.CellError, len(result.Errors))
			copy(sorted, result.Errors)
			sort.Slice(sorted, func(i, j int) bool {
				return sorted[i].Address < sorted[j].Address
			})
			for _, e := range sorted {
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

	if len(result.Errors) > 0 {
		return &ExitError{Code: 2}
	}
	return nil
}
