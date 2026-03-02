package cmd

import (
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	readPages   string
	readSlides  string
	readOffset  int
	readLimit   int
	readOutline bool
	readJSON    bool
)

var readCmd = &cobra.Command{
	Use:   "read <file-or-url>",
	Short: "Extract text from documents (PDF, DOCX, PPTX, HTML, text)",
	Long: `Extract text content or document outline from source material.

Supported formats:
  PDF   (.pdf)    Plain text extraction via PdfPig
  Word  (.doc, .docx)  Markdown via mammoth
  PPTX  (.ppt, .pptx)  Slide text extraction
  HTML  (.html, .htm)   Markdown via readability + turndown
  Text  (.txt, .md, .csv, .json, .xml, .yaml, .toml)

Navigation:
  Use --outline to get the document structure first, then target
  specific sections with --pages, --slides, or --offset/--limit.

URL support:
  Pass an HTTP(S) URL as the argument to download and read remote
  content. Content-Type is detected from the HTTP response header.

Examples:
  witan read report.pdf
  witan read report.pdf --outline
  witan read report.pdf --pages 1-5
  witan read slides.pptx --slides 1-3
  witan read notes.docx --offset 50 --limit 100
  witan read https://example.com/report.pdf --outline
  witan read data.csv --json`,
	Args: cobra.ExactArgs(1),
	RunE: runRead,
}

func init() {
	readCmd.Flags().StringVar(&readPages, "pages", "", "PDF page range (e.g. 1-5, 1,3,5)")
	readCmd.Flags().StringVar(&readSlides, "slides", "", "Presentation slide range (e.g. 1-3)")
	readCmd.Flags().IntVar(&readOffset, "offset", 0, "Start line (1-indexed)")
	readCmd.Flags().IntVar(&readLimit, "limit", 0, "Max lines to return")
	readCmd.Flags().BoolVar(&readOutline, "outline", false, "Show document structure instead of content")
	readCmd.Flags().BoolVar(&readJSON, "json", false, "Output full JSON response")
	rootCmd.AddCommand(readCmd)
}

func runRead(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	input := args[0]

	// Resolve input: URL or local file
	filePath, cleanup, err := resolveReadInput(input)
	if err != nil {
		return err
	}
	if cleanup != nil {
		defer cleanup()
	}

	key, err := resolveAPIKey()
	if err != nil {
		return err
	}

	c := newAPIClient(key)

	// Build query params
	params := url.Values{}
	if readPages != "" {
		params.Set("pages", readPages)
	}
	if readSlides != "" {
		params.Set("slides", readSlides)
	}
	if readOffset > 0 {
		params.Set("offset", fmt.Sprintf("%d", readOffset))
	}
	if readLimit > 0 {
		params.Set("limit", fmt.Sprintf("%d", readLimit))
	}

	if readOutline {
		return runReadOutline(c, filePath, params)
	}
	return runReadContent(c, filePath, params)
}

func runReadContent(c *client.Client, filePath string, params url.Values) error {
	var result *client.ReadResponse
	var err error

	if c.Stateless {
		result, err = c.Read(filePath, params)
	} else {
		var fileId, revisionId string
		fileId, revisionId, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesRead(fileId, revisionId, params)
			if client.IsNotFound(err) {
				fileId, revisionId, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesRead(fileId, revisionId, params)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	if readJSON {
		return jsonPrint(result)
	}

	// Human-friendly output: line-numbered content to stdout
	lineCount := 0
	if result.Content != "" {
		lines := strings.Split(result.Content, "\n")
		lineCount = len(lines)
		offset := result.Metadata.Offset
		for i, line := range lines {
			fmt.Printf("%6d\t%s\n", offset+i, line)
		}
	}

	// Metadata to stderr
	meta := result.Metadata
	parts := []string{}
	if meta.TotalPages != nil {
		pagesRead := ""
		if meta.ReadPages != nil {
			pagesRead = fmt.Sprintf(", %d read", *meta.ReadPages)
		}
		parts = append(parts, fmt.Sprintf("%d pages%s", *meta.TotalPages, pagesRead))
	}
	if meta.TotalSlides != nil {
		slidesRead := ""
		if meta.ReadSlides != nil {
			slidesRead = fmt.Sprintf(", %d read", *meta.ReadSlides)
		}
		parts = append(parts, fmt.Sprintf("%d slides%s", *meta.TotalSlides, slidesRead))
	}
	parts = append(parts, fmt.Sprintf("%d lines total", meta.TotalLines))
	if lineCount > 0 {
		parts = append(parts, fmt.Sprintf("showing %d–%d", meta.Offset, meta.Offset+lineCount-1))
	}
	fmt.Fprintf(os.Stderr, "%s  [%s]\n", result.Format, strings.Join(parts, ", "))

	return nil
}

func runReadOutline(c *client.Client, filePath string, params url.Values) error {
	var result *client.ReadOutlineResponse
	var err error

	if c.Stateless {
		result, err = c.ReadOutline(filePath, params)
	} else {
		var fileId, revisionId string
		fileId, revisionId, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesReadOutline(fileId, revisionId, params)
			if client.IsNotFound(err) {
				fileId, revisionId, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesReadOutline(fileId, revisionId, params)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	if readJSON {
		return jsonPrint(result)
	}

	// Human-friendly outline output
	if len(result.Outline) == 0 {
		fmt.Println("(no outline)")
	} else {
		for _, entry := range result.Outline {
			indent := strings.Repeat("  ", entry.Level)
			ref := ""
			if entry.Pages != "" {
				ref = fmt.Sprintf("  [pages %s]", entry.Pages)
			} else if entry.Slides != "" {
				ref = fmt.Sprintf("  [slide %s]", entry.Slides)
			} else if entry.Offset != nil {
				ref = fmt.Sprintf("  [line %d]", *entry.Offset)
			}
			fmt.Printf("%s%s%s\n", indent, entry.Title, ref)
		}
	}

	// Metadata to stderr
	meta := result.Metadata
	parts := []string{}
	if meta.TotalPages != nil {
		parts = append(parts, fmt.Sprintf("%d pages", *meta.TotalPages))
	}
	if meta.TotalSlides != nil {
		parts = append(parts, fmt.Sprintf("%d slides", *meta.TotalSlides))
	}
	if meta.TotalLines != nil {
		parts = append(parts, fmt.Sprintf("%d lines", *meta.TotalLines))
	}
	if len(parts) > 0 {
		fmt.Fprintf(os.Stderr, "[%s]\n", strings.Join(parts, ", "))
	}

	return nil
}

// resolveReadInput handles both local files and URLs.
// Returns the local file path and an optional cleanup function.
func resolveReadInput(input string) (string, func(), error) {
	if !strings.HasPrefix(input, "http://") && !strings.HasPrefix(input, "https://") {
		// Local file
		if _, err := os.Stat(input); err != nil {
			return "", nil, fmt.Errorf("cannot access file: %w", err)
		}
		return input, nil, nil
	}

	// URL: download to temp file
	httpClient := &http.Client{Timeout: 60 * time.Second}
	req, err := http.NewRequest("GET", input, nil)
	if err != nil {
		return "", nil, fmt.Errorf("invalid URL: %w", err)
	}
	setCLIUserAgent(req)

	resp, err := httpClient.Do(req)
	if err != nil {
		return "", nil, fmt.Errorf("downloading URL: %w", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != 200 {
		return "", nil, fmt.Errorf("downloading URL: HTTP %d", resp.StatusCode)
	}

	// Determine extension from Content-Type header, then URL path
	ext := extFromContentType(resp.Header.Get("Content-Type"))
	if ext == "" {
		ext = filepath.Ext(urlPath(input))
	}
	if ext == "" {
		ext = ".bin"
	}

	tmpFile, err := os.CreateTemp("", "witan-read-*"+ext)
	if err != nil {
		return "", nil, fmt.Errorf("creating temp file: %w", err)
	}

	if _, err := io.Copy(tmpFile, resp.Body); err != nil {
		tmpFile.Close()
		os.Remove(tmpFile.Name())
		return "", nil, fmt.Errorf("downloading URL: %w", err)
	}
	tmpFile.Close()

	cleanup := func() {
		os.Remove(tmpFile.Name())
	}
	return tmpFile.Name(), cleanup, nil
}

func extFromContentType(ct string) string {
	ct = strings.SplitN(ct, ";", 2)[0]
	ct = strings.TrimSpace(strings.ToLower(ct))
	switch ct {
	case "application/pdf":
		return ".pdf"
	case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
		return ".docx"
	case "application/msword":
		return ".doc"
	case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
		return ".pptx"
	case "application/vnd.ms-powerpoint":
		return ".ppt"
	case "text/html":
		return ".html"
	case "text/markdown":
		return ".md"
	case "text/csv":
		return ".csv"
	case "application/json":
		return ".json"
	case "application/xml", "text/xml":
		return ".xml"
	default:
		if strings.HasPrefix(ct, "text/") {
			return ".txt"
		}
		return ""
	}
}

func urlPath(rawURL string) string {
	u, err := url.Parse(rawURL)
	if err != nil {
		return ""
	}
	return u.Path
}
