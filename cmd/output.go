package cmd

import (
	"encoding/json"
	"io"
	"os"
)

// ExitError signals a non-zero exit code without printing an error message.
type ExitError struct{ Code int }

func (e *ExitError) Error() string { return "" }

func jsonPrint(v any) error {
	return jsonPrintTo(os.Stdout, v)
}

// jsonlPrint writes v as a single compact JSON line (newline-terminated) to
// stdout. Multiple calls produce newline-delimited JSON (JSONL) that a consumer
// can decode one line at a time.
func jsonlPrint(v any) error {
	return json.NewEncoder(os.Stdout).Encode(v)
}

func jsonPrintTo(w io.Writer, v any) error {
	enc := json.NewEncoder(w)
	enc.SetIndent("", "  ")
	return enc.Encode(v)
}
