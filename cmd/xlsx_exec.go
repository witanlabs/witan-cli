package cmd

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	execCode           string
	execScript         string
	execStdin          bool
	execExpr           string
	execInputJSON      string
	execTimeoutMS      int
	execMaxOutputChars int
	execSave           bool
)

var xlsxExecCmd = &cobra.Command{
	Use:   "exec <file>",
	Short: "Execute JavaScript against a workbook",
	Long: `Execute JavaScript against a workbook.

Contract:
  - Provide exactly one code source: --code, --script, --stdin, or --expr.
  - --expr wraps input as: return (<expr>);
  - --expr is for single expressions only (no semicolons/newlines); use --code for multi-statement scripts.
  - Script code must evaluate to JSON-serializable result values.

Inputs:
  - <file> is the workbook to execute against.
  - --input-json passes any JSON value to the script as input.
  - If --input-json is omitted, input defaults to {}.

Defaults:
  - --timeout-ms=0 means no explicit timeout override.
  - --max-output-chars=0 means no explicit stdout cap override.
  - --save=false means no workbook write-back.

Output:
  - Default mode prints stdout first, then:
      - pretty JSON result when ok=true
      - formatted error summary when ok=false
  - --json prints the full response envelope.
    Success shape:
      {"ok":true,"stdout":"...","result":<json>,"writes_detected":<bool>,"accesses":[...]}
      {"ok":true,...,"file":"<base64>"} when --save in stateless mode and writes are detected
      {"ok":true,...,"revision_id":"<id>"} when --save in files-backed mode and writes are detected
    Failure shape:
      {"ok":false,"stdout":"...","error":{"type":"...","code":"...","message":"..."}}

Behavior:
  - Works in both stateless and files-backed modes.
  - By default, does not overwrite the local workbook.
  - With --save, writes updated workbook bytes only when exec reports writes and the API returns file/revision output.

Exit codes:
  - 0: response has ok=true
  - 1: transport/API error, invalid request, or response has ok=false

Examples:
  witan xlsx exec report.xlsx --expr 'wb.sheet("Summary").cell("A1").value'
  witan xlsx exec report.xlsx --script ./exec.js --input-json '{"threshold":10}'
  witan xlsx exec report.xlsx --code 'console.log("hi"); return {"ok":true}'
  cat script.js | witan xlsx exec report.xlsx --stdin`,
	Args: cobra.ExactArgs(1),
	RunE: runExec,
}

func init() {
	xlsxExecCmd.Flags().StringVar(&execCode, "code", "", "Inline JavaScript source")
	xlsxExecCmd.Flags().StringVar(&execScript, "script", "", "Path to a JavaScript file")
	xlsxExecCmd.Flags().BoolVar(&execStdin, "stdin", false, "Read JavaScript source from stdin")
	xlsxExecCmd.Flags().StringVar(&execExpr, "expr", "", `Single-expression shorthand; wraps as return (<expr>);`)
	xlsxExecCmd.Flags().StringVar(&execInputJSON, "input-json", "", "JSON value passed as input to the script")
	xlsxExecCmd.Flags().IntVar(&execTimeoutMS, "timeout-ms", 0, "Execution timeout in milliseconds (> 0)")
	xlsxExecCmd.Flags().IntVar(&execMaxOutputChars, "max-output-chars", 0, "Maximum stdout characters to capture (> 0)")
	xlsxExecCmd.Flags().BoolVar(&execSave, "save", false, "Persist exec writes and overwrite local workbook when writes are detected")
	xlsxCmd.AddCommand(xlsxExecCmd)
}

func runExec(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true

	filePath, err := fixExcelExtension(args[0])
	if err != nil {
		return err
	}

	code, err := resolveExecCodeSource(cmd, os.Stdin)
	if err != nil {
		return err
	}
	if strings.TrimSpace(code) == "" {
		return fmt.Errorf("exec code must not be empty")
	}

	input, err := parseExecInput(execInputJSON, cmd.Flags().Changed("input-json"))
	if err != nil {
		return err
	}

	if err := validateExecPositiveFlag(cmd, "timeout-ms", execTimeoutMS); err != nil {
		return err
	}
	if err := validateExecPositiveFlag(cmd, "max-output-chars", execMaxOutputChars); err != nil {
		return err
	}

	req := client.ExecRequest{
		Code:           code,
		Input:          input,
		TimeoutMS:      execTimeoutMS,
		MaxOutputChars: execMaxOutputChars,
	}

	key, err := resolveAPIKey()
	if err != nil {
		return err
	}

	c := newAPIClient(key)

	var result *client.ExecResponse
	var fileID string
	if c.Stateless {
		result, err = c.Exec(filePath, req, execSave)
	} else {
		var revisionID string
		fileID, revisionID, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesExec(fileID, revisionID, req, execSave)
			if client.IsNotFound(err) {
				fileID, revisionID, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesExec(fileID, revisionID, req, execSave)
				}
			}
		}
	}
	if err != nil {
		return err
	}

	if execSave && result.Ok {
		if c.Stateless && result.File != nil {
			decoded, err := base64.StdEncoding.DecodeString(*result.File)
			if err != nil {
				return fmt.Errorf("decoding updated file: %w", err)
			}
			if err := os.WriteFile(filePath, decoded, 0o644); err != nil {
				return fmt.Errorf("writing updated file: %w", err)
			}
			if _, err := fixWritebackExtension(filePath); err != nil {
				return err
			}
		} else if !c.Stateless && result.RevisionID != nil {
			fileBytes, err := c.DownloadFileContent(fileID, *result.RevisionID)
			if err != nil {
				return fmt.Errorf("downloading updated file: %w", err)
			}
			if err := os.WriteFile(filePath, fileBytes, 0o644); err != nil {
				return fmt.Errorf("writing updated file: %w", err)
			}
			if _, err := fixWritebackExtension(filePath); err != nil {
				return err
			}
			if err := c.UpdateCachedRevision(filePath, fileID, *result.RevisionID); err != nil {
				return fmt.Errorf("updating local cache: %w", err)
			}
		}
	}

	if jsonOutput {
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
			ext := execImageExt(img)
			b64 := img
			if _, after, ok := strings.Cut(img, ","); ok {
				b64 = after
			}
			decoded, err := base64.StdEncoding.DecodeString(b64)
			if err != nil {
				return fmt.Errorf("decoding exec image: %w", err)
			}
			f, err := os.CreateTemp("", "witan-exec-*"+ext)
			if err != nil {
				return fmt.Errorf("creating temp image file: %w", err)
			}
			tmpPath := f.Name()
			f.Close()
			if err := os.WriteFile(tmpPath, decoded, 0o644); err != nil {
				os.Remove(tmpPath)
				return fmt.Errorf("writing exec image: %w", err)
			}
			fmt.Println(tmpPath)
		}
	}

	if !result.Ok {
		return &ExitError{Code: 1}
	}
	return nil
}

func resolveExecCodeSource(cmd *cobra.Command, stdin io.Reader) (string, error) {
	codeSet := cmd.Flags().Changed("code")
	scriptSet := cmd.Flags().Changed("script")
	stdinSet := execStdin
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
		if err := validateExecExpr(execExpr); err != nil {
			return "", err
		}
		return fmt.Sprintf("return (%s);", execExpr), nil
	case codeSet:
		return execCode, nil
	case scriptSet:
		if strings.TrimSpace(execScript) == "" {
			return "", fmt.Errorf("--script requires a path")
		}
		b, err := os.ReadFile(execScript)
		if err != nil {
			return "", fmt.Errorf("reading script file: %w", err)
		}
		return string(b), nil
	case stdinSet:
		b, err := io.ReadAll(stdin)
		if err != nil {
			return "", fmt.Errorf("reading --stdin: %w", err)
		}
		return string(b), nil
	default:
		return "", fmt.Errorf("exactly one of --code, --script, --stdin, or --expr is required")
	}
}

func validateExecExpr(expr string) error {
	trimmed := strings.TrimSpace(expr)
	if trimmed == "" {
		return fmt.Errorf("--expr must not be empty")
	}
	if strings.Contains(trimmed, ";") || strings.Contains(trimmed, "\n") || strings.Contains(trimmed, "\r") {
		return fmt.Errorf("--expr is for single expressions; use --code for multi-statement scripts")
	}
	return nil
}

func parseExecInput(raw string, provided bool) (any, error) {
	if !provided {
		return map[string]any{}, nil
	}
	var input any
	if err := json.Unmarshal([]byte(raw), &input); err != nil {
		return nil, fmt.Errorf("invalid --input-json: %w", err)
	}
	return input, nil
}

func validateExecPositiveFlag(cmd *cobra.Command, name string, value int) error {
	if cmd.Flags().Changed(name) && value <= 0 {
		return fmt.Errorf("--%s must be > 0", name)
	}
	return nil
}

func printExecResult(raw json.RawMessage) error {
	if len(bytes.TrimSpace(raw)) == 0 {
		return nil
	}
	var v any
	if err := json.Unmarshal(raw, &v); err != nil {
		return fmt.Errorf("parsing exec result JSON: %w", err)
	}
	return jsonPrint(v)
}

func formatExecError(execErr *client.ExecError) string {
	if execErr == nil {
		return "execution failed"
	}
	if execErr.Type != "" && execErr.Code != "" {
		return fmt.Sprintf("%s (%s): %s", execErr.Type, execErr.Code, execErr.Message)
	}
	if execErr.Code != "" {
		return fmt.Sprintf("%s: %s", execErr.Code, execErr.Message)
	}
	if execErr.Message != "" {
		return execErr.Message
	}
	return "execution failed"
}

func execImageExt(dataURL string) string {
	prefix, _, ok := strings.Cut(dataURL, ",")
	if !ok {
		return ".png"
	}
	if strings.Contains(prefix, "image/webp") {
		return ".webp"
	}
	if strings.Contains(prefix, "image/jpeg") {
		return ".jpg"
	}
	return ".png"
}
