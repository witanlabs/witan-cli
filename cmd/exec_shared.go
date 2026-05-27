package cmd

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

// resolveExecCodeSource resolves the JavaScript code to execute from various sources.
// It supports --code, --script, --stdin, and --expr flags.
// The cmd parameter is used to check which flags were set.
// The stdin parameter is used for reading from stdin (pass os.Stdin in production).
// The values are the flag values (code, script, stdinFlag, expr, stdinTimeoutMS).
func resolveExecCodeSource(cmd *cobra.Command, stdin io.Reader, code, script string, stdinFlag bool, expr string, stdinTimeoutMS int) (string, error) {
	codeSet := cmd.Flags().Changed("code")
	scriptSet := cmd.Flags().Changed("script")
	stdinSet := stdinFlag
	exprSet := cmd.Flags().Changed("expr")

	selected := 0
	for _, set := range []bool{codeSet, scriptSet, stdinSet, exprSet} {
		if set {
			selected++
		}
	}
	if selected == 0 {
		return "", fmt.Errorf("provide exactly one code source: --code, --script, --stdin, or --expr")
	}
	if selected > 1 {
		return "", fmt.Errorf("provide exactly one code source: --code, --script, --stdin, or --expr")
	}

	switch {
	case exprSet:
		if err := validateExecExpr(expr); err != nil {
			return "", err
		}
		return fmt.Sprintf("return (%s);", expr), nil
	case codeSet:
		return code, nil
	case scriptSet:
		if strings.TrimSpace(script) == "" {
			return "", fmt.Errorf("--script requires a path")
		}
		b, err := os.ReadFile(script)
		if err != nil {
			return "", fmt.Errorf("reading script file: %w", err)
		}
		return string(b), nil
	case stdinSet:
		b, err := readExecStdinWithTimeout(stdin, stdinTimeoutMS)
		if err != nil {
			return "", fmt.Errorf("reading --stdin: %w", err)
		}
		return string(b), nil
	default:
		return "", fmt.Errorf("provide exactly one code source: --code, --script, --stdin, or --expr")
	}
}

// readExecStdinWithTimeout reads from stdin with an optional timeout.
// If timeoutMS is 0, it reads without a timeout.
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
		return nil, fmt.Errorf("timed out waiting for stdin EOF after %dms (use --stdin-timeout-ms=0 to disable)", timeoutMS)
	}
}

// validateExecExpr validates that an --expr value is a single expression.
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

// outputExecResult handles the output of an exec response.
// It prints stdout, then either the result (if ok=true) or an error (if ok=false).
// If useJSON is true, it prints the full JSON response.
// If not, it prints stdout first, then pretty-prints the result or formats the error.
// Images are decoded from base64 data URLs and written to temp files.
func outputExecResult(result *client.ExecResponse, useJSON bool, formatError func(*client.ExecError) string) error {
	if useJSON {
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
			fmt.Println(formatError(result.Error))
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

// execImageExt extracts the file extension from a data URL.
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

// printExecResult pretty-prints the result JSON.
func printExecResult(raw json.RawMessage) error {
	if len(strings.TrimSpace(string(raw))) == 0 {
		return nil
	}
	var v any
	if err := json.Unmarshal(raw, &v); err != nil {
		return fmt.Errorf("parsing exec result JSON: %w", err)
	}
	return jsonPrint(v)
}

// formatExecError formats an ExecError for display.
// This is the default formatter; commands can override if they need custom error messages.
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

// validateExecPositiveFlag validates that a flag value is > 0 when explicitly set.
func validateExecPositiveFlag(cmd *cobra.Command, name string, value int) error {
	if cmd.Flags().Changed(name) && value <= 0 {
		return fmt.Errorf("--%s must be > 0", name)
	}
	return nil
}

// validateExecNonNegativeFlag validates that a flag value is >= 0 when explicitly set.
func validateExecNonNegativeFlag(cmd *cobra.Command, name string, value int) error {
	if cmd.Flags().Changed(name) && value < 0 {
		return fmt.Errorf("--%s must be >= 0", name)
	}
	return nil
}
