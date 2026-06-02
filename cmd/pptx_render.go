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
	pptxRenderSlide  int
	pptxRenderDPR    int
	pptxRenderOutput string
	pptxRenderDiff   string
)

var pptxRenderCmd = &cobra.Command{
	Use:   "render <file.pptx>",
	Short: "Render a PPTX slide as a PNG image",
	Long: `Render a PPTX slide as a PNG image.

Examples:
  witan pptx render deck.pptx --slide 1
  witan pptx render deck.pptx --slide 3 --dpr 2 -o slide-3.png
  witan pptx render deck.pptx --slide 1 --diff baseline.png`,
	Args: cobra.ExactArgs(1),
	RunE: runPPTXRender,
}

func init() {
	pptxRenderCmd.Flags().IntVar(&pptxRenderSlide, "slide", 0, "1-based slide number to render (required)")
	pptxRenderCmd.Flags().IntVar(&pptxRenderDPR, "dpr", 1, "Device pixel ratio 1-3")
	pptxRenderCmd.Flags().StringVarP(&pptxRenderOutput, "output", "o", "", "Write image to this path (default: temporary file)")
	pptxRenderCmd.Flags().StringVar(&pptxRenderDiff, "diff", "", "Compare against baseline PNG and write highlighted PNG diff")
	pptxCmd.AddCommand(pptxRenderCmd)
}

func runPPTXRender(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

	if strings.ToLower(filepath.Ext(filePath)) != ".pptx" {
		return fmt.Errorf("PPTX path must end in .pptx")
	}
	if pptxRenderSlide < 1 {
		return fmt.Errorf("--slide is required and must be >= 1")
	}
	if pptxRenderDPR < 1 || pptxRenderDPR > 3 {
		return fmt.Errorf("--dpr must be 1-3, got %d", pptxRenderDPR)
	}

	key, orgID, err := resolveAuth()
	if err != nil {
		return err
	}
	c := newAPIClient(key, orgID)

	params := map[string]string{
		"slide": strconv.Itoa(pptxRenderSlide),
		"dpr":   strconv.Itoa(pptxRenderDPR),
	}

	var imageBytes []byte
	var contentType string
	if c.Stateless {
		imageBytes, contentType, err = c.PPTXRender(filePath, params)
	} else {
		var fileID, revisionID string
		fileID, revisionID, err = c.EnsureUploaded(filePath)
		if err == nil {
			imageBytes, contentType, err = c.FilesPPTXRender(fileID, revisionID, params)
			if client.IsNotFound(err) {
				fileID, revisionID, err = c.ReuploadFile(filePath)
				if err == nil {
					imageBytes, contentType, err = c.FilesPPTXRender(fileID, revisionID, params)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	var diffSummary string
	if pptxRenderDiff != "" {
		beforeBytes, err := os.ReadFile(pptxRenderDiff)
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
		contentType = "image/png"
	}

	outPath := pptxRenderOutput
	if outPath == "" {
		f, err := os.CreateTemp("", "witan-pptx-render-*.png")
		if err != nil {
			return fmt.Errorf("creating temp file: %w", err)
		}
		outPath = f.Name()
		f.Close()
	}
	if dir := filepath.Dir(outPath); dir != "" {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return fmt.Errorf("creating output directory: %w", err)
		}
	}
	if err := os.WriteFile(outPath, imageBytes, 0o644); err != nil {
		return fmt.Errorf("writing output: %w", err)
	}

	if diffSummary != "" {
		fmt.Printf("%s\nslide=%d | dpr=%d | %s\n", outPath, pptxRenderSlide, pptxRenderDPR, diffSummary)
	} else {
		fmt.Printf("%s\nslide=%d | dpr=%d | %s\n", outPath, pptxRenderSlide, pptxRenderDPR, contentType)
	}
	return nil
}
