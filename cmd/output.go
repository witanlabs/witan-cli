package cmd

import (
	"encoding/json"
	"os"
)

// ExitError signals a non-zero exit code without printing an error message.
type ExitError struct{ Code int }

func (e *ExitError) Error() string { return "" }

func jsonPrint(v any) error {
	enc := json.NewEncoder(os.Stdout)
	enc.SetIndent("", "  ")
	return enc.Encode(v)
}
