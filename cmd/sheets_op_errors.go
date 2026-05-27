package cmd

import (
	"fmt"
	"os"

	"github.com/witanlabs/witan-cli/client"
)

// sheetsOpAuthError is the agent-readable JSON shape emitted when a Google
// Sheets operation fails because authorization is required.
type sheetsOpAuthError struct {
	Ok    bool `json:"ok"`
	Error struct {
		Code          string `json:"code"`
		SpreadsheetID string `json:"spreadsheet_id,omitempty"`
		Message       string `json:"message,omitempty"`
		Hint          string `json:"hint"`
	} `json:"error"`
}

// emitSheetsAuthError prints an authorization error (JSON or human) and returns
// an exit-code-3 ExitError so agents can branch deterministically without
// parsing messages.
func emitSheetsAuthError(code, spreadsheetID, message string, jsonOut bool) error {
	ref := "<spreadsheet>"
	if spreadsheetID != "" {
		ref = "gs://" + spreadsheetID
	}

	var hint, human string
	switch code {
	case "needs_file_authorization":
		hint = fmt.Sprintf("run 'witan gsheets authorize %s'", ref)
		human = "This Google Sheet must be authorized before Witan can access it (it may also not exist)."
	case "google_auth_required":
		hint = "run 'witan gsheets connect' to reconnect"
		human = "Google Sheets requires authorization."
	default:
		hint = fmt.Sprintf("run 'witan gsheets authorize %s'", ref)
		human = message
	}

	if jsonOut {
		var out sheetsOpAuthError
		out.Error.Code = code
		out.Error.SpreadsheetID = spreadsheetID
		out.Error.Message = message
		out.Error.Hint = hint
		if err := jsonPrint(out); err != nil {
			return err
		}
	} else {
		fmt.Fprintln(os.Stderr, human)
		fmt.Fprintln(os.Stderr, "Hint:", hint)
	}
	return &ExitError{Code: authRequiredExitCode}
}

// handleSheetsOpError converts an authorization failure from a REST sheet
// operation (exec/lint/render) into agent-friendly output and exit code 3.
// Any other error passes through unchanged.
func handleSheetsOpError(err error, spreadsheetID string, jsonOut bool) error {
	if err == nil {
		return nil
	}
	apiErr, ok := err.(*client.APIError)
	if !ok {
		return err
	}
	switch apiErr.Code {
	case "needs_file_authorization":
		return emitSheetsAuthError("needs_file_authorization", spreadsheetID, apiErr.Message, jsonOut)
	case "google_auth_required":
		return emitSheetsAuthError("google_auth_required", spreadsheetID, apiErr.Message, jsonOut)
	default:
		// Anything else (including a bare 401 with code "unauthorized", which
		// means the Witan session expired -> 'witan auth login', not reconnect)
		// falls through to the client's friendlyErrorMessage.
		return err
	}
}
