package client

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"math/rand"
	"mime/multipart"
	"net"
	"net/http"
	"net/textproto"
	"net/url"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/witanlabs/witan-cli/config"
)

const (
	defaultRequestTimeout = 60 * time.Second
	defaultMaxAttempts    = 3
	defaultBaseBackoff    = 200 * time.Millisecond
	defaultMaxBackoff     = 2 * time.Second
	defaultUserAgent      = "witan-cli/dev"
)

// Client is a Witan API client
type Client struct {
	BaseURL    string
	APIKey     string
	OrgID      string
	UserAgent  string
	HTTPClient *http.Client
	Stateless  bool       // when true, use POST-file-in-body endpoints only
	cache      *FileCache // nil when stateless

	requestTimeout time.Duration
	maxAttempts    int
	baseBackoff    time.Duration
	maxBackoff     time.Duration
	sleep          func(time.Duration)
	randInt63n     func(int64) int64
	now            func() time.Time
}

type rawResponse struct {
	StatusCode  int
	ContentType string
	RetryAfter  string
	Body        []byte
}

// New creates a new Witan API client. By default it uses the /v0/files
// endpoints with a local hash cache for deduplication. Pass stateless=true
// to use POST-file-in-body endpoints instead (zero data retention).
func New(baseURL, apiKey, orgID string, stateless bool) *Client {
	c := &Client{
		BaseURL:        strings.TrimRight(baseURL, "/"),
		APIKey:         apiKey,
		OrgID:          orgID,
		UserAgent:      defaultUserAgent,
		HTTPClient:     &http.Client{},
		Stateless:      stateless,
		requestTimeout: defaultRequestTimeout,
		maxAttempts:    defaultMaxAttempts,
		baseBackoff:    defaultBaseBackoff,
		maxBackoff:     defaultMaxBackoff,
		sleep:          time.Sleep,
		randInt63n:     rand.Int63n,
		now:            time.Now,
	}
	if !stateless {
		c.cache = NewFileCache()
		c.HTTPClient.Jar = newDefaultPersistentCookieJar()
	}
	return c
}

func newDefaultPersistentCookieJar() http.CookieJar {
	path, err := config.CookieJarPath()
	if err != nil {
		return nil
	}
	jar, err := NewPersistentCookieJar(path)
	if err != nil {
		return nil
	}
	return jar
}

// buildPath constructs an API path, inserting /orgs/{orgID} when OrgID is set.
func (c *Client) buildPath(version, path string) string {
	if c.OrgID != "" {
		return "/" + version + "/orgs/" + c.OrgID + path
	}
	return "/" + version + path
}

func (c *Client) doWithRetry(makeRequest func() (*http.Request, error)) (*rawResponse, error) {
	maxAttempts := c.maxAttempts
	if maxAttempts < 1 {
		maxAttempts = 1
	}

	for attempt := 1; attempt <= maxAttempts; attempt++ {
		req, err := makeRequest()
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}

		timeout := c.requestTimeout
		if timeout <= 0 {
			timeout = defaultRequestTimeout
		}
		ctx, cancel := context.WithTimeout(context.Background(), timeout)
		req = req.WithContext(ctx)

		resp, err := c.HTTPClient.Do(req)
		if err != nil {
			cancel()
			if attempt < maxAttempts && isRetryableTransportError(err) {
				c.sleepWithBackoff(attempt, "")
				continue
			}
			return nil, fmt.Errorf("API request failed after %d attempt(s): %w", attempt, err)
		}

		body, readErr := io.ReadAll(resp.Body)
		resp.Body.Close()
		cancel()
		if readErr != nil {
			if attempt < maxAttempts && isRetryableTransportError(readErr) {
				c.sleepWithBackoff(attempt, "")
				continue
			}
			return nil, fmt.Errorf("reading response after %d attempt(s): %w", attempt, readErr)
		}

		if attempt < maxAttempts && shouldRetryStatus(resp.StatusCode) {
			c.sleepWithBackoff(attempt, resp.Header.Get("Retry-After"))
			continue
		}

		return &rawResponse{
			StatusCode:  resp.StatusCode,
			ContentType: resp.Header.Get("Content-Type"),
			RetryAfter:  resp.Header.Get("Retry-After"),
			Body:        body,
		}, nil
	}

	return nil, fmt.Errorf("API request failed after %d attempt(s)", maxAttempts)
}

func isRetryableTransportError(err error) bool {
	if err == nil {
		return false
	}
	if errors.Is(err, context.DeadlineExceeded) {
		return true
	}
	if errors.Is(err, context.Canceled) {
		return false
	}
	var netErr net.Error
	if errors.As(err, &netErr) {
		return netErr.Timeout()
	}
	return false
}

func shouldRetryStatus(status int) bool {
	switch status {
	case http.StatusRequestTimeout, http.StatusTooManyRequests, http.StatusInternalServerError,
		http.StatusBadGateway, http.StatusServiceUnavailable, http.StatusGatewayTimeout:
		return true
	default:
		return false
	}
}

func (c *Client) sleepWithBackoff(attempt int, retryAfterHeader string) {
	if d, ok := c.parseRetryAfter(retryAfterHeader); ok {
		c.sleep(d)
		return
	}

	base := c.baseBackoff
	if base <= 0 {
		base = defaultBaseBackoff
	}
	delay := base
	for i := 1; i < attempt; i++ {
		delay *= 2
		if delay <= 0 {
			delay = defaultMaxBackoff
			break
		}
	}

	maxBackoff := c.maxBackoff
	if maxBackoff <= 0 {
		maxBackoff = defaultMaxBackoff
	}
	if delay > maxBackoff {
		delay = maxBackoff
	}
	if delay <= 0 {
		return
	}

	// Full jitter in [0, delay).
	if c.randInt63n != nil {
		delay = time.Duration(c.randInt63n(int64(delay)))
	}
	c.sleep(delay)
}

func (c *Client) parseRetryAfter(headerValue string) (time.Duration, bool) {
	v := strings.TrimSpace(headerValue)
	if v == "" {
		return 0, false
	}
	if secs, err := strconv.Atoi(v); err == nil {
		if secs <= 0 {
			return 0, false
		}
		return time.Duration(secs) * time.Second, true
	}
	if t, err := http.ParseTime(v); err == nil {
		now := time.Now
		if c.now != nil {
			now = c.now
		}
		d := t.Sub(now())
		if d > 0 {
			return d, true
		}
	}
	return 0, false
}

// Render renders a region of a spreadsheet and returns the image bytes
func (c *Client) Render(filePath string, params map[string]string) ([]byte, string, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		f, err := os.Open(filePath)
		if err != nil {
			return nil, fmt.Errorf("cannot open file: %w", err)
		}

		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/xlsx/render"))
		if err != nil {
			f.Close()
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := u.Query()
		for k, v := range params {
			q.Set(k, v)
		}
		u.RawQuery = q.Encode()

		req, err := http.NewRequest("POST", u.String(), f)
		if err != nil {
			f.Close()
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.GetBody = func() (io.ReadCloser, error) {
			return os.Open(filePath)
		}
		req.Header.Set("Content-Type", detectContentType(filePath))
		c.setCommonHeaders(req)
		return req, nil
	})
	if err != nil {
		return nil, "", err
	}

	if raw.StatusCode != 200 {
		return nil, "", parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}
	return raw.Body, raw.ContentType, nil
}

// Lint runs lint on a file via POST /v0/xlsx/lint and returns diagnostics
func (c *Client) Lint(filePath string, params url.Values) (*LintResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		f, err := os.Open(filePath)
		if err != nil {
			return nil, fmt.Errorf("cannot open file: %w", err)
		}

		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/xlsx/lint"))
		if err != nil {
			f.Close()
			return nil, fmt.Errorf("building URL: %w", err)
		}
		u.RawQuery = params.Encode()

		req, err := http.NewRequest("POST", u.String(), f)
		if err != nil {
			f.Close()
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.GetBody = func() (io.ReadCloser, error) {
			return os.Open(filePath)
		}
		req.Header.Set("Content-Type", detectContentType(filePath))
		c.setCommonHeaders(req)
		return req, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result LintResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing lint response: %w", err)
	}
	return &result, nil
}

// Calc recalculates formulas via POST /v0/xlsx/calc and returns results
func (c *Client) Calc(filePath string, params url.Values) (*CalcResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		f, err := os.Open(filePath)
		if err != nil {
			return nil, fmt.Errorf("cannot open file: %w", err)
		}

		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/xlsx/calc"))
		if err != nil {
			f.Close()
			return nil, fmt.Errorf("building URL: %w", err)
		}
		u.RawQuery = params.Encode()

		req, err := http.NewRequest("POST", u.String(), f)
		if err != nil {
			f.Close()
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.GetBody = func() (io.ReadCloser, error) {
			return os.Open(filePath)
		}
		req.Header.Set("Content-Type", detectContentType(filePath))
		c.setCommonHeaders(req)
		return req, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result CalcResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing calc response: %w", err)
	}
	return &result, nil
}

// Exec runs JavaScript against a workbook via multipart POST /v0/xlsx/exec.
func (c *Client) Exec(filePath string, req ExecRequest, save bool) (*ExecResponse, error) {
	payload, contentType, err := buildExecMultipartPayload(filePath, req, true)
	if err != nil {
		return nil, err
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/xlsx/exec"))
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := u.Query()
		if save {
			q.Set("save", "true")
		}
		if req.Locale != "" {
			q.Set("locale", req.Locale)
		}
		u.RawQuery = q.Encode()

		httpReq, err := http.NewRequest("POST", u.String(), bytes.NewReader(payload))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		httpReq.Header.Set("Content-Type", contentType)
		c.setCommonHeaders(httpReq)
		if req.Locale != "" {
			httpReq.Header.Set("Accept-Language", req.Locale)
		}
		return httpReq, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result ExecResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing exec response: %w", err)
	}
	return &result, nil
}

// ExecCreate runs JavaScript against a new workbook via multipart POST /v0/xlsx/exec?create=true.
func (c *Client) ExecCreate(filePath string, req ExecRequest, save bool) (*ExecResponse, error) {
	if req.Filename == "" {
		req.Filename = filepath.Base(filePath)
	}
	payload, contentType, err := buildExecMultipartPayload(filePath, req, false)
	if err != nil {
		return nil, err
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/xlsx/exec"))
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := u.Query()
		q.Set("create", "true")
		if save {
			q.Set("save", "true")
		}
		if req.Locale != "" {
			q.Set("locale", req.Locale)
		}
		u.RawQuery = q.Encode()

		httpReq, err := http.NewRequest("POST", u.String(), bytes.NewReader(payload))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		httpReq.Header.Set("Content-Type", contentType)
		c.setCommonHeaders(httpReq)
		if req.Locale != "" {
			httpReq.Header.Set("Accept-Language", req.Locale)
		}
		return httpReq, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result ExecResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing exec response: %w", err)
	}
	return &result, nil
}

func buildExecMultipartPayload(filePath string, req ExecRequest, includeFile bool) ([]byte, string, error) {
	var buf bytes.Buffer
	writer := multipart.NewWriter(&buf)

	if includeFile {
		f, err := os.Open(filePath)
		if err != nil {
			return nil, "", fmt.Errorf("cannot open file: %w", err)
		}
		defer f.Close()

		filename := filepath.Base(filePath)
		mimeType := detectContentType(filePath)
		h := make(textproto.MIMEHeader)
		h.Set("Content-Disposition", fmt.Sprintf(`form-data; name="file"; filename="%s"`, filename))
		h.Set("Content-Type", mimeType)
		part, err := writer.CreatePart(h)
		if err != nil {
			return nil, "", fmt.Errorf("creating form file: %w", err)
		}
		if _, err := io.Copy(part, f); err != nil {
			return nil, "", fmt.Errorf("writing file to form: %w", err)
		}
	}

	reqJSON, err := json.Marshal(req)
	if err != nil {
		return nil, "", fmt.Errorf("marshaling exec request: %w", err)
	}
	if err := writer.WriteField("exec", string(reqJSON)); err != nil {
		return nil, "", fmt.Errorf("writing exec field: %w", err)
	}

	if err := writer.Close(); err != nil {
		return nil, "", fmt.Errorf("finalizing multipart payload: %w", err)
	}

	return buf.Bytes(), writer.FormDataContentType(), nil
}

// APIError is a typed error returned by API calls, with the HTTP status code.
type APIError struct {
	StatusCode int
	Code       string
	Message    string
	RetryAfter string
}

func (e *APIError) Error() string {
	if friendly := friendlyErrorMessage(e.StatusCode, e.Code, e.Message, e.RetryAfter); friendly != "" {
		return friendly
	}
	if e.Code != "" {
		return fmt.Sprintf("API error %d: %s — %s", e.StatusCode, e.Code, e.Message)
	}
	return fmt.Sprintf("API error %d: %s", e.StatusCode, e.Message)
}

// friendlyErrorMessage translates known API error codes into user-facing messages.
func friendlyErrorMessage(statusCode int, code, message, retryAfter string) string {
	if statusCode == http.StatusTooManyRequests {
		if retryAfter != "" {
			return fmt.Sprintf("rate limited by API; retry after %s", retryAfter)
		}
		return "rate limited by API; retry in a moment"
	}
	if statusCode == http.StatusNotFound && code == "not_found" {
		if strings.Contains(message, "/pptx/") || strings.Contains(message, "/pptx") {
			return "PPTX commands are not enabled on this Witan deployment. Contact your administrator."
		}
		if strings.Contains(message, "/xlsx/") || strings.Contains(message, "/xlsx") {
			return "XLSX commands are not enabled on this Witan deployment. Contact your administrator."
		}
	}

	switch code {
	case "spawn_failed":
		return "file is not a valid Excel file (.xlsx, .xls, or .xlsm)"
	case "NOT_FOUND", "SHEET_NOT_FOUND":
		return message // already human-readable, e.g. "Sheet 'FakeSheet' not found"
	case "INVALID_ARG":
		return message
	case "ADDRESS_PARSE_ERROR":
		return message
	case "invalid_mime_type":
		if strings.Contains(strings.ToLower(message), "pptx") {
			return "unsupported file type - expected .pptx"
		}
		return "unsupported file type — expected .xlsx, .xls, or .xlsm"
	case "google_auth_required":
		return "Google Sheets requires authorization. Run 'witan gsheets connect' to enable access."
	case "google_sheets_not_found":
		return "Google Sheets spreadsheet not found or not shared with your account"
	case "google_sheets_forbidden":
		return "you don't have permission to access this Google Sheets spreadsheet"
	case "NOT_IMPLEMENTED":
		if message != "" {
			return message + " (this lint rule is not supported for Google Sheets)"
		}
		return "lint rule not supported for Google Sheets"
	case "unauthorized":
		if statusCode == 401 {
			return "authentication required — run 'witan auth login' or set WITAN_API_KEY"
		}
		return message
	default:
		return ""
	}
}

// IsNotFound returns true if the error is a 404 APIError.
func IsNotFound(err error) bool {
	if apiErr, ok := err.(*APIError); ok {
		return apiErr.StatusCode == 404 && !isRouteNotFound(apiErr)
	}
	return false
}

func isRouteNotFound(apiErr *APIError) bool {
	if apiErr == nil {
		return false
	}
	return apiErr.Code == "not_found" && strings.HasPrefix(apiErr.Message, "Route ")
}

// IsGoogleAuthRequired returns true if the error indicates Google Sheets
// authorization is needed. This happens when using a gs:// URL without
// having connected Google Sheets via 'witan gsheets connect'.
func IsGoogleAuthRequired(err error) bool {
	if apiErr, ok := err.(*APIError); ok {
		return apiErr.Code == "google_auth_required"
	}
	return false
}

func parseAPIError(statusCode int, body []byte, retryAfter string) error {
	var apiErr ErrorResponse
	if json.Unmarshal(body, &apiErr) == nil && apiErr.Error.Message != "" {
		return &APIError{
			StatusCode: statusCode,
			Code:       apiErr.Error.Code,
			Message:    apiErr.Error.Message,
			RetryAfter: retryAfter,
		}
	}
	return &APIError{StatusCode: statusCode, Message: string(body), RetryAfter: retryAfter}
}

func detectContentType(filePath string) string {
	if mt := knownMIMEType(filepath.Ext(filePath)); mt != "" {
		return mt
	}
	return "application/octet-stream"
}

// DetectContentType returns the MIME type the CLI sends for a local file.
func DetectContentType(filePath string) string {
	return detectContentType(filePath)
}

// doJSONRequest is a generic helper for making JSON requests.
// It marshals the request body (if non-nil), sends it to the given URL with the given method,
// and unmarshals the response into result (if non-nil).
// Returns the response status code or an error.
func (c *Client) doJSONRequest(method, urlPath string, reqBody, result any) error {
	var bodyBytes []byte
	var err error
	if reqBody != nil {
		bodyBytes, err = json.Marshal(reqBody)
		if err != nil {
			return fmt.Errorf("marshaling request: %w", err)
		}
	}

	resp, err := c.doWithRetry(func() (*http.Request, error) {
		var body io.Reader
		if bodyBytes != nil {
			body = bytes.NewReader(bodyBytes)
		}
		r, err := http.NewRequest(method, c.BaseURL+urlPath, body)
		if err != nil {
			return nil, err
		}
		if bodyBytes != nil {
			r.Header.Set("Content-Type", "application/json")
		}
		c.setCommonHeaders(r)
		return r, nil
	})
	if err != nil {
		return err
	}
	if resp.StatusCode != 200 && resp.StatusCode != 201 {
		return parseAPIError(resp.StatusCode, resp.Body, resp.RetryAfter)
	}

	if result != nil {
		if err := json.Unmarshal(resp.Body, result); err != nil {
			return fmt.Errorf("parsing response: %w", err)
		}
	}
	return nil
}

func (c *Client) setCommonHeaders(req *http.Request) {
	userAgent := strings.TrimSpace(c.UserAgent)
	if userAgent == "" {
		userAgent = defaultUserAgent
	}
	req.Header.Set("User-Agent", userAgent)

	if c.APIKey == "" {
		return
	}
	req.Header.Set("Authorization", "Bearer "+c.APIKey)
}

// IsGoogleSheetsURL returns true if the path looks like a Google Sheets URL.
// Supported formats:
//   - gs://SHEET_ID
//   - https://docs.google.com/spreadsheets/d/SHEET_ID/...
func IsGoogleSheetsURL(path string) bool {
	if strings.HasPrefix(path, "gs://") {
		return true
	}
	if strings.Contains(path, "docs.google.com/spreadsheets/d/") {
		return true
	}
	return false
}

// NormalizeGoogleSheetsURL converts various Google Sheets URL formats to gs://SHEET_ID.
func NormalizeGoogleSheetsURL(path string) string {
	if strings.HasPrefix(path, "gs://") {
		return path
	}
	// Extract sheet ID from web URL: https://docs.google.com/spreadsheets/d/SHEET_ID/...
	if idx := strings.Index(path, "docs.google.com/spreadsheets/d/"); idx != -1 {
		rest := path[idx+len("docs.google.com/spreadsheets/d/"):]
		// Sheet ID ends at next / or end of string
		if slashIdx := strings.Index(rest, "/"); slashIdx != -1 {
			return "gs://" + rest[:slashIdx]
		}
		// Remove any query string or fragment
		if qIdx := strings.Index(rest, "?"); qIdx != -1 {
			rest = rest[:qIdx]
		}
		if hIdx := strings.Index(rest, "#"); hIdx != -1 {
			rest = rest[:hIdx]
		}
		return "gs://" + rest
	}
	return path
}

// ExtractSpreadsheetID extracts the spreadsheet ID from a gs://SHEET_ID URL.
func ExtractSpreadsheetID(gsURL string) string {
	normalized := NormalizeGoogleSheetsURL(gsURL)
	if strings.HasPrefix(normalized, "gs://") {
		return normalized[5:]
	}
	return normalized
}

// GSheetsExec executes JavaScript against a Google Sheets spreadsheet.
// Endpoint: POST /v0/orgs/:org_id/gsheets/:spreadsheet_id/exec
func (c *Client) GSheetsExec(spreadsheetID string, req ExecRequest) (*ExecResponse, error) {
	apiPath, err := c.buildGSheetsPath(spreadsheetID, "/exec")
	if err != nil {
		return nil, err
	}

	var result ExecResponse
	if err := c.doJSONRequest("POST", apiPath, req, &result); err != nil {
		return nil, err
	}
	return &result, nil
}

// GSheetsExecCreate creates a new Google Sheet and executes JavaScript against it.
// Endpoint: POST /v0/orgs/:org_id/gsheets/new/exec
func (c *Client) GSheetsExecCreate(req ExecRequest) (*ExecResponse, error) {
	return c.GSheetsExec("new", req)
}

// buildGSheetsPath constructs an API path for Google Sheets operations.
// Unlike xlsx endpoints which work with API keys (no org) or JWTs (with org),
// Google Sheets only works with JWTs which always have an org context.
func (c *Client) buildGSheetsPath(spreadsheetID, suffix string) (string, error) {
	if c.OrgID == "" {
		return "", fmt.Errorf("Google Sheets operations require an organization context")
	}
	return "/v0/orgs/" + c.OrgID + "/gsheets/" + spreadsheetID + suffix, nil
}


// GSheetsLint runs lint diagnostics on a Google Sheets spreadsheet.
// Endpoint: GET /v0/orgs/:org_id/gsheets/:spreadsheet_id/lint
func (c *Client) GSheetsLint(spreadsheetID string, params url.Values) (*LintResponse, error) {
	apiPath, err := c.buildGSheetsPath(spreadsheetID, "/lint")
	if err != nil {
		return nil, err
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + apiPath)
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		if len(params) > 0 {
			u.RawQuery = params.Encode()
		}

		req, err := http.NewRequest("GET", u.String(), nil)
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		c.setCommonHeaders(req)
		return req, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result LintResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing lint response: %w", err)
	}
	return &result, nil
}

// GSheetsRPCWebSocketURL returns the WebSocket URL for Google Sheets RPC.
// Endpoint: GET /v0/orgs/:org_id/gsheets/ws
// Spreadsheet binding happens in the init message (spreadsheet_id or create: true).
func (c *Client) GSheetsRPCWebSocketURL() (string, error) {
	if c.OrgID == "" {
		return "", fmt.Errorf("Google Sheets operations require an organization context")
	}
	apiPath := "/v0/orgs/" + c.OrgID + "/gsheets/ws"

	wsURL := c.BaseURL + apiPath
	wsURL = strings.Replace(wsURL, "https://", "wss://", 1)
	wsURL = strings.Replace(wsURL, "http://", "ws://", 1)

	return wsURL, nil
}

// GSheetsRender renders a range of a Google Sheets spreadsheet as an image.
// Endpoint: GET /v0/orgs/:org_id/gsheets/:spreadsheet_id/render
func (c *Client) GSheetsRender(spreadsheetID string, params map[string]string) ([]byte, string, error) {
	apiPath, err := c.buildGSheetsPath(spreadsheetID, "/render")
	if err != nil {
		return nil, "", err
	}

	// Build query string
	query := url.Values{}
	for k, v := range params {
		query.Set(k, v)
	}
	fullURL := c.BaseURL + apiPath
	if len(query) > 0 {
		fullURL += "?" + query.Encode()
	}

	resp, err := c.doWithRetry(func() (*http.Request, error) {
		r, err := http.NewRequest("GET", fullURL, nil)
		if err != nil {
			return nil, err
		}
		c.setCommonHeaders(r)
		return r, nil
	})
	if err != nil {
		return nil, "", err
	}
	if resp.StatusCode != 200 {
		return nil, "", parseAPIError(resp.StatusCode, resp.Body, resp.RetryAfter)
	}

	return resp.Body, resp.ContentType, nil
}

// CreateGoogleSheetRequest is the request body for creating a new Google Sheet.
type CreateGoogleSheetRequest struct {
	Title string `json:"title,omitempty"`
}

// CreateGoogleSheetResponse is the response from creating a new Google Sheet.
type CreateGoogleSheetResponse struct {
	SpreadsheetID string `json:"spreadsheet_id"`
	Title         string `json:"title"`
	URL           string `json:"url"`
}

// CreateGoogleSheet creates a new Google Sheet in the user's Google Drive.
// Endpoint: POST /v0/orgs/:org_id/gsheets
func (c *Client) CreateGoogleSheet(title string) (*CreateGoogleSheetResponse, error) {
	if c.OrgID == "" {
		return nil, fmt.Errorf("Google Sheets operations require an organization context")
	}

	apiPath := c.buildPath("v0", "/gsheets")

	var reqBody any
	if title != "" {
		reqBody = CreateGoogleSheetRequest{Title: title}
	}

	var result CreateGoogleSheetResponse
	if err := c.doJSONRequest("POST", apiPath, reqBody, &result); err != nil {
		return nil, err
	}
	return &result, nil
}
