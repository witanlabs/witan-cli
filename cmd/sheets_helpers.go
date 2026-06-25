package cmd

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
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

// validateSheetsRef validates a reference to an EXISTING spreadsheet (used by
// lint/render/status/authorize). It does not accept the create sentinels
// new / gs://new — those are handled only by the open-or-create arg validator
// for commands that support creation (exec/rpc).
func validateSheetsRef(ref string) error {
	// Reject the create sentinel in any form (new, gs://new) — IsGoogleSheetsURL
	// accepts gs://new as a valid gs:// ref, but "new" is not an existing sheet.
	if isSheetsCreateRef(ref) {
		return fmt.Errorf("invalid spreadsheet reference: %s\n'new' creates a sheet and is not valid here; provide an existing Google Sheets URL or gs://SPREADSHEET_ID", ref)
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

// errSheetsAuthUnavailable indicates the per-file authorization state could
// not be determined (transient: quota / 5xx / network). Pollers should keep
// retrying; one-shot callers should report "unknown" rather than "not
// authorized".
var errSheetsAuthUnavailable = errors.New("could not determine sheet authorization status; please retry")

type authorizeSheetStartRequest struct {
	Spreadsheet string `json:"spreadsheet"`
	RedirectURL string `json:"redirect_url,omitempty"`
}

type authorizeSheetStartResponse struct {
	PickerURL string `json:"picker_url"`
}

type authorizeSheetStatusResponse struct {
	Authorized bool `json:"authorized"`
}

// authorizeSheetStart begins per-file authorization for a spreadsheet and
// returns the Google Picker URL to open in a browser. spreadsheet may be an
// id, gs:// reference, or docs URL (canonicalized server-side).
func authorizeSheetStart(httpClient httpDoer, mgmtURL, jwt, spreadsheet, redirectURL string) (string, error) {
	body, _ := json.Marshal(authorizeSheetStartRequest{Spreadsheet: spreadsheet, RedirectURL: redirectURL})
	req, err := http.NewRequest("POST", mgmtURL+"/v0/integrations/google-sheets/authorize-sheet/start", bytes.NewReader(body))
	if err != nil {
		return "", fmt.Errorf("failed to create authorize-sheet request: %w", err)
	}
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Authorization", "Bearer "+jwt)
	setCLIUserAgent(req)

	resp, err := httpClient.Do(req)
	if err != nil {
		return "", fmt.Errorf("failed to start sheet authorization: %w", err)
	}
	defer resp.Body.Close()

	respBody, _ := io.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK {
		return "", sheetsAuthorizeError(parseManagementAPIError(resp.StatusCode, respBody))
	}

	var out authorizeSheetStartResponse
	if err := json.Unmarshal(respBody, &out); err != nil {
		return "", fmt.Errorf("failed to parse authorize-sheet response: %w", err)
	}
	if out.PickerURL == "" {
		return "", fmt.Errorf("authorize-sheet response missing picker_url")
	}
	return out.PickerURL, nil
}

// authorizeSheetStatus reports whether a spreadsheet is authorized for the
// app. Transient conditions (rate limiting, upstream 5xx, or network/transport
// errors) are returned as errSheetsAuthUnavailable so callers can distinguish
// "not authorized" from "couldn't determine" and keep polling.
func authorizeSheetStatus(httpClient httpDoer, mgmtURL, jwt, spreadsheet string) (bool, error) {
	u := mgmtURL + "/v0/integrations/google-sheets/authorize-sheet/status?spreadsheet=" + url.QueryEscape(spreadsheet)
	req, err := http.NewRequest("GET", u, nil)
	if err != nil {
		return false, fmt.Errorf("failed to create authorize-sheet status request: %w", err)
	}
	req.Header.Set("Authorization", "Bearer "+jwt)
	setCLIUserAgent(req)

	resp, err := httpClient.Do(req)
	if err != nil {
		// Network/transport blip: transient, worth retrying during a poll.
		return false, fmt.Errorf("%w: %v", errSheetsAuthUnavailable, err)
	}
	defer resp.Body.Close()

	respBody, _ := io.ReadAll(resp.Body)

	if resp.StatusCode != http.StatusOK {
		if isTransientStatusCode(resp.StatusCode) {
			return false, errSheetsAuthUnavailable
		}
		return false, sheetsAuthorizeError(parseManagementAPIError(resp.StatusCode, respBody))
	}

	var out authorizeSheetStatusResponse
	if err := json.Unmarshal(respBody, &out); err != nil {
		return false, fmt.Errorf("failed to parse authorize-sheet status: %w", err)
	}
	return out.Authorized, nil
}

// isTransientStatusCode reports whether an HTTP status indicates a transient
// condition worth retrying during a poll (rate limiting or upstream 5xx).
func isTransientStatusCode(code int) bool {
	switch code {
	case http.StatusTooManyRequests,
		http.StatusInternalServerError,
		http.StatusBadGateway,
		http.StatusServiceUnavailable,
		http.StatusGatewayTimeout:
		return true
	}
	return false
}

// isTransientManagementError reports whether an error from a management API
// status call is transient (worth continuing to poll rather than aborting).
// Transport/network and parse errors (which aren't typed *ManagementAPIError)
// are treated as transient; typed errors are transient only for rate-limit/5xx
// status codes. Auth errors (401/403/404) are not transient.
func isTransientManagementError(err error) bool {
	apiErr, ok := err.(*ManagementAPIError)
	if !ok {
		return true
	}
	return isTransientStatusCode(apiErr.StatusCode)
}

// sheetsAuthorizeError translates management API error codes from the
// authorize-sheet endpoints into actionable user-facing messages.
func sheetsAuthorizeError(err error) error {
	apiErr, ok := err.(*ManagementAPIError)
	if !ok {
		return err
	}
	switch apiErr.Code {
	case "google_sheets_not_connected", "google_sheets_scope_not_granted":
		return fmt.Errorf("Google Sheets is not connected. Run 'witan gsheets connect' first.")
	case "google_auth_required":
		return fmt.Errorf("Google authorization expired or was revoked. Run 'witan gsheets connect' to reconnect.")
	case "forbidden":
		return fmt.Errorf("Google Sheets requires user authentication.\nRun 'witan auth login' and try again without --api-key or WITAN_API_KEY.")
	case "unauthorized":
		return fmt.Errorf("session expired: run 'witan auth login' to re-authenticate")
	case "bad_request":
		if apiErr.Message != "" {
			return fmt.Errorf("could not parse spreadsheet reference: %s", apiErr.Message)
		}
		return fmt.Errorf("could not parse spreadsheet reference")
	default:
		return err
	}
}
