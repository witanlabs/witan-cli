package cmd

import (
	"fmt"
	"strconv"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
	"github.com/witanlabs/witan-cli/internal"
)

var (
	renderRange  string
	renderDPR    int
	renderFormat string
	renderOutput string
	renderDiff   string
)

var renderCmd = &cobra.Command{
	Use:   "render <file>",
	Short: "Render a sheet range as an image",
	Long: `Render a sheet-qualified range as a PNG or WebP image.

Behavior:
  - --range is required (for example "Sheet1!A1:Z50").
  - --format supports png or webp.
  - --dpr must be 1-3; default is auto.
  - If --output is omitted, the image is written to a temporary file.
  - --diff compares against a baseline PNG and writes a highlighted PNG diff.
  - Large images (>1568 px in either dimension) may be downscaled by vision models.

Examples:
  witan xlsx render report.xlsx -r "Sheet1!A1:Z50"
  witan xlsx render report.xlsx -r "'My Sheet'!B5:H20" --dpr 2
  witan xlsx render report.xlsx -r "Sheet1!A1:F10" -o before.png
  witan xlsx render report.xlsx -r "Sheet1!A1:F10" --diff before.png`,
	Args: cobra.ExactArgs(1),
	RunE: runRender,
}

func init() {
	renderCmd.Flags().StringVarP(&renderRange, "range", "r", "", `Sheet-qualified range to render (required)`)
	renderCmd.Flags().IntVar(&renderDPR, "dpr", 0, "Device pixel ratio 1-3 (default: auto)")
	renderCmd.Flags().StringVar(&renderFormat, "format", "png", "Output image format: png or webp")
	renderCmd.Flags().StringVarP(&renderOutput, "output", "o", "", "Write image to this path (default: temporary file)")
	renderCmd.Flags().StringVar(&renderDiff, "diff", "", "Compare against baseline PNG and write highlighted diff image")
	xlsxCmd.AddCommand(renderCmd)
}

func runRender(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

	filePath, err := fixExcelExtension(filePath)
	if err != nil {
		return err
	}

	key, orgID, err := resolveAuth()
	if err != nil {
		return err
	}

	if renderFormat != "png" && renderFormat != "webp" {
		return fmt.Errorf("--format must be 'png' or 'webp', got %q", renderFormat)
	}

	c := newAPIClient(key, orgID)

	// Require --range (syntax is server-validated)
	if renderRange == "" {
		return fmt.Errorf("--range is required (e.g. -r \"Sheet1!A1:Z50\" or \"'My Sheet'!A1:Z50\")")
	}

	address := renderRange

	// Auto DPR heuristic
	dpr := renderDPR
	if dpr == 0 {
		dpr = autoDPR(address)
	}
	if dpr < 1 || dpr > 3 {
		return fmt.Errorf("--dpr must be 1-3, got %d", dpr)
	}

	// Render
	params := map[string]string{
		"address": address,
		"dpr":     strconv.Itoa(dpr),
		"format":  renderFormat,
	}

	var imageBytes []byte
	var contentType string
	if c.Stateless {
		imageBytes, contentType, err = c.Render(filePath, params)
	} else {
		var fileId, revisionId string
		fileId, revisionId, err = c.EnsureUploaded(filePath)
		if err == nil {
			imageBytes, contentType, err = c.FilesRender(fileId, revisionId, params)
			if client.IsNotFound(err) {
				fileId, revisionId, err = c.ReuploadFile(filePath)
				if err == nil {
					imageBytes, contentType, err = c.FilesRender(fileId, revisionId, params)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	// If --diff is set, pixel-diff against the baseline image
	var diffSummary string
	if renderDiff != "" {
		var err error
		imageBytes, diffSummary, err = runRenderDiffPipeline(renderFormat, renderDiff, imageBytes)
		if err != nil {
			return err
		}
		contentType = "image/png"
	}

	// Write image
	outPath, err := writeRenderedImage(renderOutput, contentType, imageBytes)
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

