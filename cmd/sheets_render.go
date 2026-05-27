package cmd

import (
	"fmt"
	"strconv"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
	"github.com/witanlabs/witan-cli/internal"
)

var (
	sheetsRenderRange  string
	sheetsRenderDPR    int
	sheetsRenderFormat string
	sheetsRenderOutput string
	sheetsRenderDiff   string
)

var sheetsRenderCmd = &cobra.Command{
	Use:   "render <spreadsheet>",
	Short: "Render a sheet range as an image",
	Long: `Render a range from a Google Sheets spreadsheet as a PNG or WebP image.

Spreadsheet reference:
  You can reference spreadsheets using either format:
  - Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  - Short form: gs://SPREADSHEET_ID

Behavior:
  - --range is required (for example "Sheet1!A1:Z50").
  - --format supports png or webp.
  - --dpr must be 1-3; default is auto.
  - If --output is omitted, the image is written to a temporary file.
  - --diff compares against a baseline PNG and writes a highlighted PNG diff.
  - Large images (>1568 px in either dimension) may be downscaled by vision models.

Examples:
  witan gsheets render gs://SPREADSHEET_ID -r "Sheet1!A1:Z50"
  witan gsheets render "https://docs.google.com/spreadsheets/d/ID/edit" -r "'My Sheet'!B5:H20" --dpr 2
  witan gsheets render gs://ID -r "Sheet1!A1:F10" -o before.png
  witan gsheets render gs://ID -r "Sheet1!A1:F10" --diff before.png`,
	Args: cobra.ExactArgs(1),
	RunE: runSheetsRender,
}

func init() {
	sheetsRenderCmd.SilenceUsage = true
	sheetsRenderCmd.Flags().StringVarP(&sheetsRenderRange, "range", "r", "", `Sheet-qualified range to render (required)`)
	sheetsRenderCmd.Flags().IntVar(&sheetsRenderDPR, "dpr", 0, "Device pixel ratio 1-3 (default: auto)")
	sheetsRenderCmd.Flags().StringVar(&sheetsRenderFormat, "format", "png", "Output image format: png or webp")
	sheetsRenderCmd.Flags().StringVarP(&sheetsRenderOutput, "output", "o", "", "Write image to this path (default: temporary file)")
	sheetsRenderCmd.Flags().StringVar(&sheetsRenderDiff, "diff", "", "Compare against baseline PNG and write highlighted diff image")
	gsheetsCmd.AddCommand(sheetsRenderCmd)
}

func runSheetsRender(cmd *cobra.Command, args []string) error {
	spreadsheetRef := args[0]

	if err := validateSheetsRef(spreadsheetRef); err != nil {
		return err
	}

	// Validate format
	if sheetsRenderFormat != "png" && sheetsRenderFormat != "webp" {
		return fmt.Errorf("--format must be 'png' or 'webp', got %q", sheetsRenderFormat)
	}

	// Require --range
	if sheetsRenderRange == "" {
		return fmt.Errorf("--range is required (e.g. -r \"Sheet1!A1:Z50\" or \"'My Sheet'!A1:Z50\")")
	}

	address := sheetsRenderRange

	// Auto DPR heuristic
	dpr := sheetsRenderDPR
	if dpr == 0 {
		dpr = autoDPR(address)
	}
	if dpr < 1 || dpr > 3 {
		return fmt.Errorf("--dpr must be 1-3, got %d", dpr)
	}

	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}

	spreadsheetID := client.ExtractSpreadsheetID(spreadsheetRef)

	// Render
	params := map[string]string{
		"address": address,
		"dpr":     strconv.Itoa(dpr),
		"format":  sheetsRenderFormat,
	}

	imageBytes, contentType, err := auth.Client.GSheetsRender(spreadsheetID, params)
	if err != nil {
		return handleSheetsOpError(err, spreadsheetID, gsheetsJSONOutput)
	}

	// If --diff is set, pixel-diff against the baseline image
	var diffSummary string
	if sheetsRenderDiff != "" {
		var err error
		imageBytes, diffSummary, err = runRenderDiffPipeline(sheetsRenderFormat, sheetsRenderDiff, imageBytes)
		if err != nil {
			return err
		}
		contentType = "image/png"
	}

	// Write image
	outPath, err := writeRenderedImage(sheetsRenderOutput, contentType, imageBytes)
	if err != nil {
		return err
	}

	// Print result info
	rangeStr := address
	pixelWidth, pixelHeight := 0, 0
	if sheet, sr, sc, er, ec, parseErr := internal.ParseRange(address); parseErr == nil {
		rangeStr = internal.FormatAddress(sheet, sr, sc, er, ec)
		pixelWidth, pixelHeight = estimatePixels(address, dpr)
	}

	printRenderResult(outPath, rangeStr, pixelWidth, pixelHeight, dpr, diffSummary)
	return nil
}

