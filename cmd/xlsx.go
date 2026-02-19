package cmd

import "github.com/spf13/cobra"

var jsonOutput bool

var xlsxCmd = &cobra.Command{
	Use:   "xlsx",
	Short: "Spreadsheet commands",
	Long: `Operate on Excel workbooks (.xls, .xlsx, .xlsm).

Commands:
  calc   Recalculate formulas and update cached values in the workbook.
  edit   Update cell values, formulas, or formats and save the workbook.
  lint   Run semantic formula checks and report diagnostics.
  render Render a sheet range as PNG or WebP.

Output:
  default  Human-friendly summaries
  --json   Raw JSON responses for automation

Examples:
  witan xlsx calc report.xlsx
  witan xlsx --json lint report.xlsx
  witan xlsx render report.xlsx -r "Sheet1!A1:F20"`,
}

func init() {
	xlsxCmd.PersistentFlags().BoolVar(&jsonOutput, "json", false, "Output raw JSON instead of human-formatted summaries")
	rootCmd.AddCommand(xlsxCmd)
}
