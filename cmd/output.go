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

func jsonPrintTo(w io.Writer, v any) error {
	enc := json.NewEncoder(w)
	enc.SetIndent("", "  ")
	return enc.Encode(v)
}
