package cmd

import (
	"encoding/base64"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	pptxExecCode           string
	pptxExecScript         string
	pptxExecStdin          bool
	pptxExecExpr           string
	pptxExecInputJSON      string
	pptxExecLocale         string
	pptxExecStdinTimeoutMS int
	pptxExecTimeoutMS      int
	pptxExecMaxOutputChars int
	pptxExecSave           bool
	pptxExecCreate         bool
)

var pptxExecCmd = &cobra.Command{
	Use:   "exec <file.pptx>",
	Short: "Execute Office.js-compatible JavaScript against a PPTX file",
	Long: `Execute Office.js-compatible JavaScript against a PPTX file.

Provide exactly one code source: --code, --script, --stdin, or --expr.
Use --create with a new .pptx path to start from an empty PPTX file.
Use --save to write changes back to the local file.

Examples:
  witan pptx exec deck.pptx --expr 'PowerPoint.run(async context => { const count = context.presentation.slides.getCount(); await context.sync(); return count.value })'
  witan pptx exec deck.pptx --stdin --save < edit.js
  witan pptx exec new.pptx --create --save --code 'return await PowerPoint.run(async context => { const slides = context.presentation.slides; slides.add(); const count = slides.getCount(); await context.sync(); return count.value })'`,
	Args: cobra.ExactArgs(1),
	RunE: runPPTXExec,
}

func init() {
	pptxExecCmd.Flags().StringVar(&pptxExecCode, "code", "", "Inline JavaScript source")
	pptxExecCmd.Flags().StringVar(&pptxExecScript, "script", "", "Path to a JavaScript file")
	pptxExecCmd.Flags().BoolVar(&pptxExecStdin, "stdin", false, "Read JavaScript source from stdin")
	pptxExecCmd.Flags().StringVar(&pptxExecExpr, "expr", "", "Single-expression shorthand; wraps as return (<expr>);")
	pptxExecCmd.Flags().StringVar(&pptxExecInputJSON, "input-json", "", "JSON value passed as input to the script")
	pptxExecCmd.Flags().StringVar(&pptxExecLocale, "locale", "", "Execution locale (env: WITAN_LOCALE; otherwise LC_ALL / LC_MESSAGES / LANG)")
	pptxExecCmd.Flags().IntVar(&pptxExecStdinTimeoutMS, "stdin-timeout-ms", defaultExecStdinTimeoutMS, "Maximum time to wait for EOF when reading --stdin (0 disables)")
	pptxExecCmd.Flags().IntVar(&pptxExecTimeoutMS, "timeout-ms", 0, "Execution timeout in milliseconds (> 0)")
	pptxExecCmd.Flags().IntVar(&pptxExecMaxOutputChars, "max-output-chars", 0, "Maximum stdout characters to capture (> 0)")
	pptxExecCmd.Flags().BoolVar(&pptxExecCreate, "create", false, "Create a new .pptx file instead of opening an existing file")
	pptxExecCmd.Flags().BoolVar(&pptxExecSave, "save", false, "Write returned PPTX bytes to the target path")
	pptxCmd.AddCommand(pptxExecCmd)
}

func runPPTXExec(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true

	filePath, err := resolvePPTXExecPresentationPath(args[0], pptxExecCreate)
	if err != nil {
		return err
	}
	if err := validateExecPositiveFlag(cmd, "timeout-ms", pptxExecTimeoutMS); err != nil {
		return err
	}
	if err := validateExecNonNegativeFlag(cmd, "stdin-timeout-ms", pptxExecStdinTimeoutMS); err != nil {
		return err
	}
	if err := validateExecPositiveFlag(cmd, "max-output-chars", pptxExecMaxOutputChars); err != nil {
		return err
	}

	code, err := resolvePPTXExecCodeSource(cmd, os.Stdin)
	if err != nil {
		return err
	}
	if strings.TrimSpace(code) == "" {
		return fmt.Errorf("exec code must not be empty")
	}

	input, err := parseExecInput(pptxExecInputJSON, cmd.Flags().Changed("input-json"))
	if err != nil {
		return err
	}

	locale, err := resolvePPTXExecLocale(cmd)
	if err != nil {
		return err
	}

	req := client.ExecRequest{
		Code:           code,
		Input:          input,
		Locale:         locale,
		TimeoutMS:      pptxExecTimeoutMS,
		MaxOutputChars: pptxExecMaxOutputChars,
	}

	key, orgID, err := resolveAuth()
	if err != nil {
		return err
	}

	c := newAPIClient(key, orgID)
	if pptxExecCreate {
		c = client.New(resolveAPIURL(), key, orgID, true)
		c.UserAgent = cliUserAgent()
	}

	var result *client.ExecResponse
	var fileID string
	if pptxExecCreate {
		result, err = c.PPTXExecCreate(filePath, req, pptxExecSave)
	} else if c.Stateless {
		result, err = c.PPTXExec(filePath, req, pptxExecSave)
	} else {
		var revisionID string
		fileID, revisionID, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesPPTXExec(fileID, revisionID, req, pptxExecSave)
			if client.IsNotFound(err) {
				fileID, revisionID, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesPPTXExec(fileID, revisionID, req, pptxExecSave)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	if pptxExecSave && result.Ok {
		if pptxExecCreate || c.Stateless {
			if result.File != nil {
				decoded, err := base64.StdEncoding.DecodeString(*result.File)
				if err != nil {
					return fmt.Errorf("decoding PPTX bytes: %w", err)
				}
				if err := os.WriteFile(filePath, decoded, 0o644); err != nil {
					return fmt.Errorf("writing PPTX file: %w", err)
				}
			} else if pptxExecCreate {
				return fmt.Errorf("creating PPTX file: expected file bytes in response")
			}
		} else if result.RevisionID != nil {
			fileBytes, err := c.DownloadFileContent(fileID, *result.RevisionID)
			if err != nil {
				return fmt.Errorf("downloading updated PPTX file: %w", err)
			}
			if err := os.WriteFile(filePath, fileBytes, 0o644); err != nil {
				return fmt.Errorf("writing updated PPTX file: %w", err)
			}
			if err := c.UpdateCachedRevision(filePath, fileID, *result.RevisionID); err != nil {
				return fmt.Errorf("updating local cache: %w", err)
			}
		}
	}

	if pptxJSONOutput {
		result.File = nil
		if err := jsonPrint(result); err != nil {
			return err
		}
	} else {
		if result.Stdout != "" {
			fmt.Print(result.Stdout)
		}
		if result.Ok {
			if err := printExecResult(result.Result); err != nil {
				return err
			}
		} else {
			fmt.Println(formatExecError(result.Error))
		}
		for _, img := range result.Images {
			if err := writePPTXExecImage(img); err != nil {
				return err
			}
		}
	}

	if !result.Ok {
		return &ExitError{Code: 1}
	}
	return nil
}

func resolvePPTXExecPresentationPath(filePath string, create bool) (string, error) {
	if strings.ToLower(filepath.Ext(filePath)) != ".pptx" {
		return "", fmt.Errorf("PPTX path must end in .pptx")
	}
	if !create {
		info, err := os.Stat(filePath)
		if err != nil {
			return "", fmt.Errorf("checking PPTX file: %w", err)
		}
		if info.IsDir() {
			return "", fmt.Errorf("PPTX path is a directory: %s", filePath)
		}
		return filePath, nil
	}

	if _, err := os.Stat(filePath); err == nil {
		return "", fmt.Errorf("--create requires a target path that does not already exist")
	} else if !errors.Is(err, os.ErrNotExist) {
		return "", fmt.Errorf("checking target path: %w", err)
	}

	parent := filepath.Dir(filePath)
	info, err := os.Stat(parent)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return "", fmt.Errorf("parent directory does not exist: %s", parent)
		}
		return "", fmt.Errorf("checking parent directory: %w", err)
	}
	if !info.IsDir() {
		return "", fmt.Errorf("parent path is not a directory: %s", parent)
	}
	return filePath, nil
}

func resolvePPTXExecCodeSource(cmd *cobra.Command, stdin io.Reader) (string, error) {
	codeSet := cmd.Flags().Changed("code")
	scriptSet := cmd.Flags().Changed("script")
	stdinSet := pptxExecStdin
	exprSet := cmd.Flags().Changed("expr")

	selected := 0
	for _, set := range []bool{codeSet, scriptSet, stdinSet, exprSet} {
		if set {
			selected++
		}
	}
	if selected == 0 {
		return "", fmt.Errorf("exactly one of --code, --script, --stdin, or --expr is required")
	}
	if selected > 1 {
		return "", fmt.Errorf("--code, --script, --stdin, and --expr are mutually exclusive")
	}

	switch {
	case exprSet:
		if err := validateExecExpr(pptxExecExpr); err != nil {
			return "", err
		}
		return fmt.Sprintf("return (%s);", pptxExecExpr), nil
	case codeSet:
		return pptxExecCode, nil
	case scriptSet:
		if strings.TrimSpace(pptxExecScript) == "" {
			return "", fmt.Errorf("--script requires a path")
		}
		b, err := os.ReadFile(pptxExecScript)
		if err != nil {
			return "", fmt.Errorf("reading script file: %w", err)
		}
		return string(b), nil
	case stdinSet:
		b, err := readExecStdinWithTimeout(stdin, pptxExecStdinTimeoutMS)
		if err != nil {
			return "", fmt.Errorf("reading --stdin: %w", err)
		}
		return string(b), nil
	default:
		return "", fmt.Errorf("exactly one of --code, --script, --stdin, or --expr is required")
	}
}

func resolvePPTXExecLocale(cmd *cobra.Command) (string, error) {
	if cmd.Flags().Changed("locale") {
		locale, ok := normalizeLocale(pptxExecLocale)
		if !ok {
			return "", fmt.Errorf("invalid --locale %q", pptxExecLocale)
		}
		return locale, nil
	}
	if raw, ok := os.LookupEnv("WITAN_LOCALE"); ok && strings.TrimSpace(raw) != "" {
		locale, valid := normalizeLocale(raw)
		if !valid {
			return "", fmt.Errorf("invalid WITAN_LOCALE %q", raw)
		}
		return locale, nil
	}
	if raw, ok := os.LookupEnv("LC_ALL"); ok && strings.TrimSpace(raw) != "" {
		locale, _ := normalizeLocale(raw)
		return locale, nil
	}
	for _, key := range []string{"LC_MESSAGES", "LANG"} {
		raw, ok := os.LookupEnv(key)
		if !ok || strings.TrimSpace(raw) == "" {
			continue
		}
		if locale, valid := normalizeLocale(raw); valid {
			return locale, nil
		}
	}
	return "", nil
}

func writePPTXExecImage(dataURL string) error {
	ext := execImageExt(dataURL)
	b64 := dataURL
	if _, after, ok := strings.Cut(dataURL, ","); ok {
		b64 = after
	}
	decoded, err := base64.StdEncoding.DecodeString(b64)
	if err != nil {
		return fmt.Errorf("decoding exec image: %w", err)
	}
	f, err := os.CreateTemp("", "witan-pptx-exec-*"+ext)
	if err != nil {
		return fmt.Errorf("creating temp image file: %w", err)
	}
	tmpPath := f.Name()
	if _, err := f.Write(decoded); err != nil {
		f.Close()
		os.Remove(tmpPath)
		return fmt.Errorf("writing exec image: %w", err)
	}
	if err := f.Close(); err != nil {
		os.Remove(tmpPath)
		return fmt.Errorf("closing exec image file: %w", err)
	}
	fmt.Println(tmpPath)
	return nil
}
