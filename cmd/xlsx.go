package cmd

import "github.com/spf13/cobra"

var jsonOutput bool

var xlsxCmd = &cobra.Command{
	Use:   "xlsx",
	Short: "Excel spreadsheet commands",
}

func init() {
	xlsxCmd.PersistentFlags().BoolVar(&jsonOutput, "json", false, "Output raw JSON instead of human-formatted text")
	rootCmd.AddCommand(xlsxCmd)
}
