package cmd

import (
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"

	"github.com/witanlabs/witan-cli/client"
	"github.com/witanlabs/witan-cli/config"
)

// sheetsAuthResult holds the result of authenticating for Google Sheets operations.
type sheetsAuthResult struct {
	Client *client.Client
	JWT    string
	OrgID  string
	MgmtURL string
}

// requireSheetsAuth validates that the user is authenticated with a session (not API key)
// and returns a configured client for Google Sheets operations.
func requireSheetsAuth() (*sheetsAuthResult, error) {
	// Require session auth (not API key)
	if resolveRawAPIKey() != "" {
		return nil, fmt.Errorf("Google Sheets requires user authentication.\nRun 'witan auth login' and try again without --api-key or WITAN_API_KEY.")
	}

	cfg, err := config.Load()
	if err != nil {
		return nil, fmt.Errorf("loading auth config: %w", err)
	}
	if cfg.SessionToken == "" {
		return nil, fmt.Errorf("not authenticated: run 'witan auth login' first")
	}

	mgmtURL := resolveManagementAPIURL()

	// Exchange session token for JWT
	jwt, err := exchangeSessionForJWT(mgmtURL, cfg.SessionToken)
	if err != nil {
		return nil, fmt.Errorf("session expired: run 'witan auth login' to re-authenticate")
	}

	return &sheetsAuthResult{
		Client:  newAPIClient(jwt, cfg.SessionOrgID),
		JWT:     jwt,
		OrgID:   cfg.SessionOrgID,
		MgmtURL: mgmtURL,
	}, nil
}

// isSheetsCreateRef reports whether ref selects create+exec mode.
func isSheetsCreateRef(ref string) bool {
	return client.ExtractSpreadsheetID(ref) == "new"
}

// validateSheetsTitle validates an optional spreadsheet title for create mode.
func validateSheetsTitle(title string, create bool) error {
	if title == "" {
		return nil
	}
	if !create {
		return nil
	}
	if len(title) > 1000 {
		return fmt.Errorf("--title must be at most 1000 characters")
	}
	return nil
}

// validateSheetsRef validates that the given reference is a valid Google Sheets URL or gs:// reference.
func validateSheetsRef(ref string) error {
	if isSheetsCreateRef(ref) {
		return nil
	}
	if !client.IsGoogleSheetsURL(ref) {
		return fmt.Errorf("invalid spreadsheet reference: %s\nExpected a Google Sheets URL or gs://SPREADSHEET_ID", ref)
	}
	return nil
}

// resolveSheetsCreate reports whether a gsheets command should open in create mode.
func resolveSheetsCreate(createFlag bool, args []string) bool {
	if createFlag {
		return true
	}
	if len(args) == 1 && isSheetsCreateRef(args[0]) {
		return true
	}
	return false
}

// validateSheetsOpenOrCreateArgs validates spreadsheet arguments for open or create mode.
func validateSheetsOpenOrCreateArgs(args []string, createFlag bool) error {
	create := resolveSheetsCreate(createFlag, args)
	if create {
		if createFlag && len(args) == 1 && !isSheetsCreateRef(args[0]) {
			return fmt.Errorf("--create requires spreadsheet reference 'new' or gs://new, or omit the argument")
		}
		if len(args) > 1 {
			return fmt.Errorf("accepts at most 1 spreadsheet reference when using --create")
		}
		return nil
	}
	if len(args) != 1 {
		return fmt.Errorf("requires exactly 1 spreadsheet reference, or use --create to make a new sheet")
	}
	return validateSheetsRef(args[0])
}

// outputSheetsCreateHints prints stderr guidance after creating a spreadsheet.
func outputSheetsCreateHints(spreadsheetID, sheetURL, title string) {
	if spreadsheetID == "" && sheetURL == "" {
		return
	}

	if title != "" {
		fmt.Fprintf(os.Stderr, "Created Google Sheet: %s\n", title)
	} else {
		fmt.Fprintln(os.Stderr, "Created Google Sheet:")
	}
	if sheetURL != "" {
		fmt.Fprintf(os.Stderr, "URL: %s\n", sheetURL)
	}
	if spreadsheetID != "" {
		fmt.Fprintf(os.Stderr, "\nUse with gsheets commands:\n")
		fmt.Fprintf(os.Stderr, "  witan gsheets exec gs://%s --expr '...'\n", spreadsheetID)
		fmt.Fprintf(os.Stderr, "  witan gsheets rpc gs://%s\n", spreadsheetID)
	}
}


// ManagementAPIError represents a structured error from the management API.
type ManagementAPIError struct {
	StatusCode int
	Code       string
	Message    string
}

func (e *ManagementAPIError) Error() string {
	if e.Code != "" {
		return fmt.Sprintf("%s: %s", e.Code, e.Message)
	}
	if e.Message != "" {
		return e.Message
	}
	return fmt.Sprintf("API error %d", e.StatusCode)
}

// parseManagementAPIError parses an error response from the management API.
// Returns a *ManagementAPIError that callers can type-assert to check the Code.
func parseManagementAPIError(statusCode int, body []byte) error {
	var errResp struct {
		Error struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"error"`
	}
	if json.Unmarshal(body, &errResp) == nil && (errResp.Error.Code != "" || errResp.Error.Message != "") {
		return &ManagementAPIError{
			StatusCode: statusCode,
			Code:       errResp.Error.Code,
			Message:    errResp.Error.Message,
		}
	}
	return &ManagementAPIError{
		StatusCode: statusCode,
		Message:    string(body),
	}
}

// sheetsIntegrationStatus is the management API Google Sheets connection status.
type sheetsIntegrationStatus struct {
	Connected bool   `json:"connected"`
	ExpiresAt string `json:"expires_at,omitempty"`
	Status    string `json:"status,omitempty"` // "active" or "needs_reauth"
}

// isActive reports whether the integration is connected and ready to use.
func (s *sheetsIntegrationStatus) isActive() bool {
	return s != nil && s.Connected && s.Status == "active"
}

// needsReauth reports whether Google credentials exist but must be refreshed.
func (s *sheetsIntegrationStatus) needsReauth() bool {
	return s != nil && s.Connected && s.Status == "needs_reauth"
}

// getGoogleSheetsIntegrationStatus fetches Google Sheets connection status from the management API.
func getGoogleSheetsIntegrationStatus(httpClient httpDoer, mgmtURL, jwt string) (*sheetsIntegrationStatus, error) {
	req, err := http.NewRequest("GET", mgmtURL+"/v0/integrations/google-sheets/status", nil)
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+jwt)
	setCLIUserAgent(req)

	resp, err := httpClient.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)

	if resp.StatusCode != http.StatusOK {
		return nil, parseManagementAPIError(resp.StatusCode, body)
	}

	var status sheetsIntegrationStatus
	if err := json.Unmarshal(body, &status); err != nil {
		return nil, fmt.Errorf("failed to parse Google Sheets status: %w", err)
	}
	return &status, nil
}

// sheetsStatusCheckError formats management API errors from status polling.
func sheetsStatusCheckError(err error) error {
	apiErr, ok := err.(*ManagementAPIError)
	if !ok {
		return fmt.Errorf("checking connection status: %w", err)
	}
	switch apiErr.Code {
	case "unauthorized":
		return fmt.Errorf("session expired: run 'witan auth login' to re-authenticate")
	case "forbidden":
		return fmt.Errorf("Google Sheets integration requires user authentication.\nRun 'witan auth login' and try again without --api-key or WITAN_API_KEY.")
	default:
		return fmt.Errorf("checking connection status: %w", err)
	}
}

// httpDoer is an interface for HTTP clients (allows testing).
type httpDoer interface {
	Do(*http.Request) (*http.Response, error)
}
