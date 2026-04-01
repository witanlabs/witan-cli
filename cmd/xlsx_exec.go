package cmd

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	execCode           string
	execScript         string
	execStdin          bool
	execExpr           string
	execInputJSON      string
	execLocale         string
	execStdinTimeoutMS int
	execTimeoutMS      int
	execMaxOutputChars int
	execSave           bool
	execCreate         bool
)

const defaultExecStdinTimeoutMS = 2000

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
  - <file> is the workbook to execute against, or the new .xlsx target path when --create is set.
  - --input-json passes any JSON value to the script as input.
  - --locale sets the workbook execution locale explicitly.
  - If --input-json is omitted, input defaults to {}.

Defaults:
  - If --locale is omitted, the CLI tries WITAN_LOCALE, then LC_ALL / LC_MESSAGES / LANG.
  - --timeout-ms=0 means no explicit timeout override.
  - --stdin-timeout-ms=2000 aborts --stdin reads that never reach EOF; set 0 to disable.
  - --max-output-chars=0 means no explicit stdout cap override.
  - --create=false means exec expects an existing workbook path.
  - --save=false means no workbook write-back.

Output:
  - Default mode prints stdout first, then:
      - pretty JSON result when ok=true
      - formatted error summary when ok=false
  - --json prints the full response envelope.
    Success shape:
      {"ok":true,"stdout":"...","result":<json>,"writes_detected":<bool>,"accesses":[...]}
      {"ok":true,...,"revision_id":"<id>"} when --save in files-backed mode and writes are detected
    Failure shape:
      {"ok":false,"stdout":"...","error":{"type":"...","code":"...","message":"..."}}

Behavior:
  - Works in both stateless and files-backed modes.
  - --create starts a new workbook instead of opening an existing file.
  - --create requires a target path ending in .xlsx that does not already exist.
  - By default, does not overwrite the local workbook.
  - With --save, writes updated workbook bytes when the API returns file/revision output.
  - With --create --save, writes the newly created workbook to the target path.

Exit codes:
  - 0: response has ok=true
  - 1: transport/API error, invalid request, or response has ok=false

Examples:
  witan xlsx exec report.xlsx --expr 'await xlsx.readCell(wb, "Summary!A1")'
  witan xlsx exec report.xlsx --script ./exec.js --input-json '{"threshold":10}'
  witan xlsx exec report.xlsx --code 'console.log("hi"); return {"ok":true}'
  witan xlsx exec model.xlsx --create --save --code 'await xlsx.addSheet(wb, "Inputs"); return true'
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
	xlsxExecCmd.Flags().StringVar(&execLocale, "locale", "", "Execution locale (env: WITAN_LOCALE; otherwise LC_ALL / LC_MESSAGES / LANG)")
	xlsxExecCmd.Flags().IntVar(&execStdinTimeoutMS, "stdin-timeout-ms", defaultExecStdinTimeoutMS, "Maximum time to wait for EOF when reading --stdin (0 disables)")
	xlsxExecCmd.Flags().IntVar(&execTimeoutMS, "timeout-ms", 0, "Execution timeout in milliseconds (> 0)")
	xlsxExecCmd.Flags().IntVar(&execMaxOutputChars, "max-output-chars", 0, "Maximum stdout characters to capture (> 0)")
	xlsxExecCmd.Flags().BoolVar(&execCreate, "create", false, "Create a new .xlsx workbook instead of opening an existing file; target path must not exist")
	xlsxExecCmd.Flags().BoolVar(&execSave, "save", false, "Write returned workbook bytes to the target path")
	xlsxCmd.AddCommand(xlsxExecCmd)
}

func runExec(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true

	filePath, err := resolveExecWorkbookPath(args[0], execCreate)
	if err != nil {
		return err
	}

	if err := validateExecPositiveFlag(cmd, "timeout-ms", execTimeoutMS); err != nil {
		return err
	}
	if err := validateExecNonNegativeFlag(cmd, "stdin-timeout-ms", execStdinTimeoutMS); err != nil {
		return err
	}
	if err := validateExecPositiveFlag(cmd, "max-output-chars", execMaxOutputChars); err != nil {
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

	locale, err := resolveExecLocale(cmd)
	if err != nil {
		return err
	}

	req := client.ExecRequest{
		Code:           code,
		Input:          input,
		Filename:       "",
		Locale:         locale,
		TimeoutMS:      execTimeoutMS,
		MaxOutputChars: execMaxOutputChars,
	}
	if execCreate {
		req.Filename = filepath.Base(filePath)
	}

	key, orgID, err := resolveAuth()
	if err != nil {
		return err
	}

	c := newAPIClient(key, orgID)
	if execCreate {
		c = client.New(resolveAPIURL(), key, orgID, true)
		c.UserAgent = cliUserAgent()
	}

	var result *client.ExecResponse
	var fileID string
	if execCreate {
		result, err = c.ExecCreate(filePath, req, execSave)
	} else if c.Stateless {
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
		if execCreate {
			if result.File == nil {
				return fmt.Errorf("creating workbook: expected file bytes in response")
			}
			decoded, err := base64.StdEncoding.DecodeString(*result.File)
			if err != nil {
				return fmt.Errorf("decoding created file: %w", err)
			}
			if err := os.WriteFile(filePath, decoded, 0o644); err != nil {
				return fmt.Errorf("writing created file: %w", err)
			}
			if _, err := fixWritebackExtension(filePath); err != nil {
				return err
			}
		} else if c.Stateless && result.File != nil {
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
			if filePath, err = fixWritebackExtension(filePath); err != nil {
				return err
			}
			if err := c.UpdateCachedRevision(filePath, fileID, *result.RevisionID); err != nil {
				return fmt.Errorf("updating local cache: %w", err)
			}
		}
	}

	if jsonOutput {
		// File is a transport detail used for local writeback, not CLI JSON output.
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
		}
	}

	if !result.Ok {
		return &ExitError{Code: 1}
	}
	return nil
}

func resolveExecWorkbookPath(filePath string, create bool) (string, error) {
	if !create {
		return fixExcelExtension(filePath)
	}

	if strings.ToLower(filepath.Ext(filePath)) != ".xlsx" {
		return "", fmt.Errorf("--create requires a target path ending in .xlsx")
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
		b, err := readExecStdinWithTimeout(stdin, execStdinTimeoutMS)
		if err != nil {
			return "", fmt.Errorf("reading --stdin: %w", err)
		}
		return string(b), nil
	default:
		return "", fmt.Errorf("exactly one of --code, --script, --stdin, or --expr is required")
	}
}

func readExecStdinWithTimeout(stdin io.Reader, timeoutMS int) ([]byte, error) {
	if timeoutMS == 0 {
		return io.ReadAll(stdin)
	}

	type readResult struct {
		b   []byte
		err error
	}
	done := make(chan readResult, 1)
	go func() {
		b, err := io.ReadAll(stdin)
		done <- readResult{b: b, err: err}
	}()

	timer := time.NewTimer(time.Duration(timeoutMS) * time.Millisecond)
	defer timer.Stop()

	select {
	case res := <-done:
		return res.b, res.err
	case <-timer.C:
		return nil, fmt.Errorf("stdin read timed out after %dms waiting for EOF; ensure the input stream closes or set --stdin-timeout-ms=0 to disable timeout", timeoutMS)
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

func resolveExecLocale(cmd *cobra.Command) (string, error) {
	if cmd.Flags().Changed("locale") {
		locale, ok := normalizeLocale(execLocale)
		if !ok {
			return "", fmt.Errorf("invalid --locale %q", execLocale)
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

func normalizeLocale(raw string) (string, bool) {
	value := strings.TrimSpace(raw)
	if value == "" {
		return "", true
	}

	if idx := strings.IndexByte(value, '.'); idx >= 0 {
		value = value[:idx]
	}
	if idx := strings.IndexByte(value, '@'); idx >= 0 {
		value = value[:idx]
	}

	value = strings.TrimSpace(strings.ReplaceAll(value, "_", "-"))
	if value == "" {
		return "", true
	}

	upper := strings.ToUpper(value)
	if upper == "C" || upper == "POSIX" || value == "*" {
		return "", false
	}

	parts := strings.Split(value, "-")
	for i, part := range parts {
		if part == "" || !isLocaleToken(part) {
			return "", false
		}
		switch {
		case i == 0:
			parts[i] = strings.ToLower(part)
		case len(part) == 4 && isAlpha(part):
			parts[i] = strings.ToUpper(part[:1]) + strings.ToLower(part[1:])
		case len(part) == 2 && isAlpha(part):
			parts[i] = strings.ToUpper(part)
		case len(part) == 3 && isNumeric(part):
			parts[i] = part
		default:
			parts[i] = strings.ToLower(part)
		}
	}

	return strings.Join(parts, "-"), true
}

func isLocaleToken(s string) bool {
	for _, r := range s {
		if (r < 'a' || r > 'z') && (r < 'A' || r > 'Z') && (r < '0' || r > '9') {
			return false
		}
	}
	return true
}

func isAlpha(s string) bool {
	for _, r := range s {
		if (r < 'a' || r > 'z') && (r < 'A' || r > 'Z') {
			return false
		}
	}
	return true
}

func isNumeric(s string) bool {
	for _, r := range s {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func validateExecPositiveFlag(cmd *cobra.Command, name string, value int) error {
	if cmd.Flags().Changed(name) && value <= 0 {
		return fmt.Errorf("--%s must be > 0", name)
	}
	return nil
}

func validateExecNonNegativeFlag(cmd *cobra.Command, name string, value int) error {
	if cmd.Flags().Changed(name) && value < 0 {
		return fmt.Errorf("--%s must be >= 0", name)
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
