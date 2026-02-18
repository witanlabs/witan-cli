package cmd

import (
	"bytes"
	"fmt"
	"image/png"
	"os"
	"path/filepath"
	"strconv"
	"strings"

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
	Short: "Render a spreadsheet region as an image",
	Long: `Render a spreadsheet region as a PNG or WebP image.

Requires --range with a sheet-qualified address (e.g. "Sheet1!A1:Z50").

Examples:
  witan xlsx render report.xlsx -r "Sheet1!A1:Z50"
  witan xlsx render report.xlsx -r "'My Sheet'!B5:H20" --dpr 2
  witan xlsx render report.xlsx -r "Sheet1!A1:F10" -o before.png   # save baseline
  witan xlsx render report.xlsx -r "Sheet1!A1:F10" --diff before.png  # highlight changes`,
	Args: cobra.ExactArgs(1),
	RunE: runRender,
}

func init() {
	renderCmd.Flags().StringVarP(&renderRange, "range", "r", "", `Sheet-qualified range (e.g. "Sheet1!A1:Z50")`)
	renderCmd.Flags().IntVar(&renderDPR, "dpr", 0, "Device pixel ratio (1-3, auto by default)")
	renderCmd.Flags().StringVar(&renderFormat, "format", "png", "Image format (png or webp)")
	renderCmd.Flags().StringVarP(&renderOutput, "output", "o", "", "Output file path (default: temp file)")
	renderCmd.Flags().StringVar(&renderDiff, "diff", "", "Path to baseline image; outputs a pixel-diff highlight instead of the raw render")
	xlsxCmd.AddCommand(renderCmd)
}

func runRender(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

	key, err := resolveAPIKey()
	if err != nil {
		return err
	}

	if renderFormat != "png" && renderFormat != "webp" {
		return fmt.Errorf("--format must be 'png' or 'webp', got %q", renderFormat)
	}

	c := client.New(resolveAPIURL(), key, resolveStateless())

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
		if renderFormat != "png" {
			return fmt.Errorf("--diff requires --format png (got %q)", renderFormat)
		}

		beforeBytes, err := os.ReadFile(renderDiff)
		if err != nil {
			return fmt.Errorf("reading baseline image: %w", err)
		}
		beforeImg, err := png.Decode(bytes.NewReader(beforeBytes))
		if err != nil {
			return fmt.Errorf("decoding baseline image: %w", err)
		}
		afterImg, err := png.Decode(bytes.NewReader(imageBytes))
		if err != nil {
			return fmt.Errorf("decoding rendered image: %w", err)
		}

		diffImg, changed, err := internal.DiffImages(beforeImg, afterImg)
		if err != nil {
			return fmt.Errorf("diffing images: %w", err)
		}

		total := diffImg.Bounds().Dx() * diffImg.Bounds().Dy()
		diffSummary = internal.FormatDiffSummary(changed, total)

		var buf bytes.Buffer
		if err := png.Encode(&buf, diffImg); err != nil {
			return fmt.Errorf("encoding diff image: %w", err)
		}
		imageBytes = buf.Bytes()
		// Force png content type for diff output
		contentType = "image/png"
	}

	// Determine output path
	outPath := renderOutput
	if outPath == "" {
		ext := ".png"
		if strings.Contains(contentType, "webp") {
			ext = ".webp"
		}
		f, err := os.CreateTemp("", "witan-render-*"+ext)
		if err != nil {
			return fmt.Errorf("creating temp file: %w", err)
		}
		outPath = f.Name()
		f.Close()
	}

	// Ensure output directory exists
	if dir := filepath.Dir(outPath); dir != "" {
		os.MkdirAll(dir, 0o755)
	}

	if err := os.WriteFile(outPath, imageBytes, 0o644); err != nil {
		return fmt.Errorf("writing output: %w", err)
	}

	// Print result info
	rangeStr := address
	pixelWidth, pixelHeight := 0, 0
	if sheet, sr, sc, er, ec, parseErr := internal.ParseRange(address); parseErr == nil {
		rangeStr = internal.FormatAddress(sheet, sr, sc, er, ec)
		pixelWidth, pixelHeight = estimatePixels(address, dpr)
	}

	if diffSummary != "" {
		if pixelWidth > 0 && pixelHeight > 0 {
			fmt.Printf("%s\n%s | ~%d×%dpx | dpr=%d | %s\n", outPath, rangeStr, pixelWidth, pixelHeight, dpr, diffSummary)
		} else {
			fmt.Printf("%s\n%s | dpr=%d | %s\n", outPath, rangeStr, dpr, diffSummary)
		}
	} else {
		if pixelWidth > 0 && pixelHeight > 0 {
			fmt.Printf("%s\n%s | ~%d×%dpx | dpr=%d\n", outPath, rangeStr, pixelWidth, pixelHeight, dpr)
		} else {
			fmt.Printf("%s\n%s | dpr=%d\n", outPath, rangeStr, dpr)
		}
	}

	// Vision model warning (on stdout — it's actionable advice for agents)
	if pixelWidth > 1568 || pixelHeight > 1568 {
		fmt.Printf("Warning: Image exceeds 1568px. Vision models may downscale, reducing detail. Consider a smaller --range.\n")
	}

	return nil
}

func autoDPR(address string) int {
	_, sr, sc, er, ec, err := internal.ParseRange(address)
	if err != nil {
		return 2 // default
	}
	cols := ec - sc + 1
	rows := er - sr + 1
	estWidth := cols * 64
	estHeight := rows * 15
	if estWidth*2 > 1568 || estHeight*2 > 1568 {
		return 1
	}
	return 2
}

func estimatePixels(address string, dpr int) (int, int) {
	_, sr, sc, er, ec, err := internal.ParseRange(address)
	if err != nil {
		return 0, 0
	}
	cols := ec - sc + 1
	rows := er - sr + 1
	return cols * 64 * dpr, rows * 15 * dpr
}
