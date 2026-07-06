package cmd

import "github.com/spf13/cobra"

var pptxJSONOutput bool

var pptxCmd = &cobra.Command{
	Use:   "pptx",
	Short: "PPTX commands",
	Long: `Operate on PPTX files (.pptx).

Commands:
  exec   Execute Office.js-compatible JavaScript against existing PPTX files or create new .pptx files with --create.
  render Render a PPTX slide as PNG.
  lint   Run semantic presentation checks.

Output:
  default  Human-friendly summaries
  --json   Raw JSON responses for automation

Examples:
  witan pptx render deck.pptx --slide 1 -o slide-1.png
  witan pptx exec deck.pptx --expr 'PowerPoint.run(async context => { const count = context.presentation.slides.getCount(); await context.sync(); return count.value })'
  witan pptx --json exec deck.pptx --expr 'PowerPoint.run(async context => { const count = context.presentation.slides.getCount(); await context.sync(); return count.value })'`,
}

func init() {
	pptxCmd.PersistentFlags().BoolVar(&pptxJSONOutput, "json", false, "Output raw JSON instead of human-formatted summaries")
	rootCmd.AddCommand(pptxCmd)
}
