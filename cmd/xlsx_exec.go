package cmd

import (
	"bytes"
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
)

var xlsxExecCmd = &cobra.Command{
	Use:   "exec <file>",
	Short: "Execute JavaScript against a workbook",
	Long: `Execute JavaScript against a workbook.

Contract:
  - Provide exactly one code source: --code, --script, --stdin, or --expr.
  - --expr wraps input as: return (<expr>);
  - Script code must evaluate to JSON-serializable result values.

Inputs:
  - <file> is the workbook to execute against.
  - --input-json passes any JSON value to the script as input.
  - If --input-json is omitted, input defaults to {}.

Defaults:
  - --timeout-ms=0 means no explicit timeout override.
  - --max-output-chars=0 means no explicit stdout cap override.

Output:
  - Default mode prints stdout first, then:
      - pretty JSON result when ok=true
      - formatted error summary when ok=false
  - --json prints the full response envelope.
    Success shape:
      {"ok":true,"stdout":"...","result":<json>,"writes_detected":<bool>,"accesses":[...]}
    Failure shape:
      {"ok":false,"stdout":"...","error":{"type":"...","code":"...","message":"..."}}

Behavior:
  - Works in both stateless and files-backed modes.
  - Never writes workbook bytes from exec responses back to disk.

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
	xlsxExecCmd.Flags().StringVar(&execExpr, "expr", "", `Expression shorthand; wraps as return (<expr>);`)
	xlsxExecCmd.Flags().StringVar(&execInputJSON, "input-json", "", "JSON value passed as input to the script")
	xlsxExecCmd.Flags().IntVar(&execTimeoutMS, "timeout-ms", 0, "Execution timeout in milliseconds (> 0)")
	xlsxExecCmd.Flags().IntVar(&execMaxOutputChars, "max-output-chars", 0, "Maximum stdout characters to capture (> 0)")
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

	c := client.New(resolveAPIURL(), key, resolveStateless())

	var result *client.ExecResponse
	if c.Stateless {
		result, err = c.Exec(filePath, req)
	} else {
		var fileID, revisionID string
		fileID, revisionID, err = c.EnsureUploaded(filePath)
		if err == nil {
			result, err = c.FilesExec(fileID, revisionID, req)
			if client.IsNotFound(err) {
				fileID, revisionID, err = c.ReuploadFile(filePath)
				if err == nil {
					result, err = c.FilesExec(fileID, revisionID, req)
				}
			}
		}
	}
	if err != nil {
		return err
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
