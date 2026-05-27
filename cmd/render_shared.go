package cmd

import (
	"bytes"
	"fmt"
	"image/png"
	"os"
	"path/filepath"
	"strings"

	"github.com/witanlabs/witan-cli/internal"
)

// autoDPR calculates an appropriate device pixel ratio based on the range size.
// It aims to keep the rendered image under 1568px in either dimension.
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

// estimatePixels estimates the pixel dimensions of a rendered range.
func estimatePixels(address string, dpr int) (int, int) {
	_, sr, sc, er, ec, err := internal.ParseRange(address)
	if err != nil {
		return 0, 0
	}
	cols := ec - sc + 1
	rows := er - sr + 1
	return cols * 64 * dpr, rows * 15 * dpr
}

// runRenderDiffPipeline compares a baseline PNG image with a new rendered image.
// It returns the diff image bytes and a formatted summary string.
// The format parameter must be "png" or this will return an error.
// The baselinePath is the path to the baseline PNG file.
// The renderedBytes are the new rendered image bytes.
func runRenderDiffPipeline(format string, baselinePath string, renderedBytes []byte) (diffBytes []byte, summary string, err error) {
	if format != "png" {
		return nil, "", fmt.Errorf("--diff requires --format png (got %q)", format)
	}

	beforeBytes, err := os.ReadFile(baselinePath)
	if err != nil {
		return nil, "", fmt.Errorf("reading baseline image: %w", err)
	}
	beforeImg, err := png.Decode(bytes.NewReader(beforeBytes))
	if err != nil {
		return nil, "", fmt.Errorf("decoding baseline image: %w", err)
	}
	afterImg, err := png.Decode(bytes.NewReader(renderedBytes))
	if err != nil {
		return nil, "", fmt.Errorf("decoding rendered image: %w", err)
	}

	diffImg, changed, err := internal.DiffImages(beforeImg, afterImg)
	if err != nil {
		return nil, "", fmt.Errorf("diffing images: %w", err)
	}

	total := diffImg.Bounds().Dx() * diffImg.Bounds().Dy()
	summary = internal.FormatDiffSummary(changed, total)

	var buf bytes.Buffer
	if err := png.Encode(&buf, diffImg); err != nil {
		return nil, "", fmt.Errorf("encoding diff image: %w", err)
	}
	return buf.Bytes(), summary, nil
}

// writeRenderedImage writes image bytes to the specified output path.
// If outPath is empty, creates a temp file with appropriate extension.
// Returns the actual path written to.
func writeRenderedImage(outPath string, contentType string, imageBytes []byte) (string, error) {
	if outPath == "" {
		ext := ".png"
		if strings.Contains(contentType, "webp") {
			ext = ".webp"
		}
		f, err := os.CreateTemp("", "witan-render-*"+ext)
		if err != nil {
			return "", fmt.Errorf("creating temp file: %w", err)
		}
		outPath = f.Name()
		f.Close()
	}

	// Ensure output directory exists
	if dir := filepath.Dir(outPath); dir != "" {
		os.MkdirAll(dir, 0o755)
	}

	if err := os.WriteFile(outPath, imageBytes, 0o644); err != nil {
		return "", fmt.Errorf("writing output: %w", err)
	}
	return outPath, nil
}

// printRenderResult prints render output info and warnings.
func printRenderResult(outPath, rangeStr string, pixelW, pixelH, dpr int, diffSummary string) {
	if diffSummary != "" {
		if pixelW > 0 && pixelH > 0 {
			fmt.Printf("%s\n%s | ~%d×%dpx | dpr=%d | %s\n", outPath, rangeStr, pixelW, pixelH, dpr, diffSummary)
		} else {
			fmt.Printf("%s\n%s | dpr=%d | %s\n", outPath, rangeStr, dpr, diffSummary)
		}
	} else {
		if pixelW > 0 && pixelH > 0 {
			fmt.Printf("%s\n%s | ~%d×%dpx | dpr=%d\n", outPath, rangeStr, pixelW, pixelH, dpr)
		} else {
			fmt.Printf("%s\n%s | dpr=%d\n", outPath, rangeStr, dpr)
		}
	}

	// Vision model warning
	if pixelW > 1568 || pixelH > 1568 {
		fmt.Printf("Warning: Image exceeds 1568px. Vision models may downscale, reducing detail. Consider a smaller --range.\n")
	}
}
