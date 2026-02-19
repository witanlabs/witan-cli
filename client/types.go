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

// EditCell is a single cell edit request
type EditCell struct {
	Address string          `json:"address"`
	Value   json.RawMessage `json:"value,omitempty"`
	Formula string          `json:"formula,omitempty"`
	Format  string          `json:"format,omitempty"`
}

// EditResponse is the response from the edit endpoint
type EditResponse struct {
	Touched          map[string]string `json:"touched"`
	Errors           []CellError       `json:"errors"`
	InvalidatedTiles json.RawMessage   `json:"invalidatedTiles"`
	UpdatedSheets    json.RawMessage   `json:"updatedSheets"`
	File             *string           `json:"file,omitempty"`
	RevisionID       *string           `json:"revision_id,omitempty"`
}
