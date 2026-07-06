package cmd

import (
	"os"

	"github.com/spf13/cobra"
)

var pptxExecTypesCmd = &cobra.Command{
	Use:   "exec-types",
	Short: "Print the TypeScript declarations for the pptx exec sandbox",
	Long: `Print the combined TypeScript declarations available to pptx exec scripts.

The output combines the stripped Office.js PowerPoint surface and the Witan
chart extensions in a single .d.ts blob (~2 MB). It requires no authentication.

Because the output is raw TypeScript, the global --json flag is ignored. Redirect
it to a file once and search with rg rather than reading it wholesale:

  witan pptx exec-types > "${TMPDIR:-/tmp}/witan-pptx-types.d.ts"
  rg -n "SlideCollection|addChart|ChartSeries" "${TMPDIR:-/tmp}/witan-pptx-types.d.ts"`,
	Args: cobra.NoArgs,
	RunE: runPPTXExecTypes,
}

func init() {
	pptxCmd.AddCommand(pptxExecTypesCmd)
}

func runPPTXExecTypes(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true

	// The endpoint is public, so skip resolveAuth() entirely and send an
	// unauthenticated request. This keeps the command working in environments
	// that have never run `witan auth login`.
	c := newAPIClient("", "")

	body, err := c.PPTXExecTypes()
	if err != nil {
		return err
	}

	if _, err := os.Stdout.Write(body); err != nil {
		return err
	}
	return nil
}
