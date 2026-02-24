package cmd

import "github.com/spf13/cobra"

var jsonOutput bool

var xlsxCmd = &cobra.Command{
	Use:   "xlsx",
	Short: "Spreadsheet commands",
	Long: `Operate on Excel workbooks (.xls, .xlsx, .xlsm).

Commands:
  calc   Recalculate formulas, update cached values, or run non-mutating verification with --verify.
  exec   Execute JavaScript with workbook read/write access (persist writes with --save).
  lint   Run semantic formula checks and report diagnostics.
  render Render a sheet range as PNG or WebP.

Output:
  default  Human-friendly summaries
  --json   Raw JSON responses for automation

Examples:
  witan xlsx calc report.xlsx
  witan xlsx exec report.xlsx --expr 'wb.sheet("Summary").cell("A1").value'
  witan xlsx --json lint report.xlsx
  witan xlsx render report.xlsx -r "Sheet1!A1:F20"`,
}

func init() {
	xlsxCmd.PersistentFlags().BoolVar(&jsonOutput, "json", false, "Output raw JSON instead of human-formatted summaries")
	rootCmd.AddCommand(xlsxCmd)
}
