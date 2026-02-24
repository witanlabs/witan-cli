package client

import "encoding/json"

// ErrorResponse is the standard API error shape
type ErrorResponse struct {
	Error struct {
		Code    string `json:"code"`
		Message string `json:"message"`
	} `json:"error"`
}

// LintDiagnostic is a single lint diagnostic
type LintDiagnostic struct {
	Severity string  `json:"severity"`
	RuleId   string  `json:"ruleId"`
	Message  string  `json:"message"`
	Location *string `json:"location"`
}

// LintResponse is the response from the lint endpoint
type LintResponse struct {
	Diagnostics []LintDiagnostic `json:"diagnostics"`
	Total       int              `json:"total"`
}

// CellError is a formula calculation error
type CellError struct {
	Address string  `json:"address"`
	Code    string  `json:"code"`
	Formula *string `json:"formula"`
	Detail  *string `json:"detail"`
}

// CalcTouchedCell is a cell that was recalculated
type CalcTouchedCell struct {
	Value   string  `json:"value"`
	Formula *string `json:"formula"`
}

// CalcResponse is the response from the calc endpoint
type CalcResponse struct {
	Touched    map[string]CalcTouchedCell `json:"touched"`
	Changed    []string                   `json:"changed,omitempty"` // cells whose computed value changed
	Errors     []CellError                `json:"errors"`
	File       *string                    `json:"file,omitempty"`        // base64, stateless only
	RevisionID *string                    `json:"revision_id,omitempty"` // new revision, files-backed only
}

// ExecRequest is the request body for exec endpoints.
type ExecRequest struct {
	Code           string `json:"code"`
	Input          any    `json:"input,omitempty"`
	TimeoutMS      int    `json:"timeout_ms,omitempty"`
	MaxOutputChars int    `json:"max_output_chars,omitempty"`
}

// ExecAccess describes a workbook access observed during execution.
type ExecAccess struct {
	Operation string `json:"operation"` // read|write
	Address   string `json:"address"`
}

// ExecError describes a script execution error.
type ExecError struct {
	Type    string `json:"type"` // syntax|runtime|timeout
	Code    string `json:"code"` // EXEC_SYNTAX_ERROR|EXEC_RUNTIME_ERROR|EXEC_TIMEOUT|EXEC_RESULT_TOO_LARGE
	Message string `json:"message"`
}

// ExecResponse is the response from exec endpoints.
type ExecResponse struct {
	Ok             bool            `json:"ok"`
	Stdout         string          `json:"stdout"`
	Truncated      bool            `json:"truncated,omitempty"`
	Result         json.RawMessage `json:"result,omitempty"`
	WritesDetected bool            `json:"writes_detected,omitempty"`
	Accesses       []ExecAccess    `json:"accesses,omitempty"`
	File           *string         `json:"file,omitempty"`        // base64, stateless save=true only
	RevisionID     *string         `json:"revision_id,omitempty"` // new revision, files-backed save=true only
	Error          *ExecError      `json:"error,omitempty"`
}
