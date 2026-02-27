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
func New(baseURL, apiKey string, stateless bool) *Client {
	c := &Client{
		BaseURL:        strings.TrimRight(baseURL, "/"),
		APIKey:         apiKey,
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
	}
	return c
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

		u, err := url.Parse(c.BaseURL + "/v0/xlsx/render")
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

		u, err := url.Parse(c.BaseURL + "/v0/xlsx/lint")
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

		u, err := url.Parse(c.BaseURL + "/v0/xlsx/calc")
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
	payload, contentType, err := buildExecMultipartPayload(filePath, req)
	if err != nil {
		return nil, err
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + "/v0/xlsx/exec")
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		if save {
			q := u.Query()
			q.Set("save", "true")
			u.RawQuery = q.Encode()
		}

		httpReq, err := http.NewRequest("POST", u.String(), bytes.NewReader(payload))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		httpReq.Header.Set("Content-Type", contentType)
		c.setCommonHeaders(httpReq)
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

func buildExecMultipartPayload(filePath string, req ExecRequest) ([]byte, string, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, "", fmt.Errorf("cannot open file: %w", err)
	}
	defer f.Close()

	var buf bytes.Buffer
	writer := multipart.NewWriter(&buf)

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
		return "unsupported file type — expected .xlsx, .xls, or .xlsm"
	default:
		return ""
	}
}

// IsNotFound returns true if the error is a 404 APIError.
func IsNotFound(err error) bool {
	if apiErr, ok := err.(*APIError); ok {
		return apiErr.StatusCode == 404
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
	lower := strings.ToLower(filePath)
	switch {
	case strings.HasSuffix(lower, ".xlsx"):
		return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	case strings.HasSuffix(lower, ".xls"):
		return "application/vnd.ms-excel"
	case strings.HasSuffix(lower, ".xlsm"):
		return "application/vnd.ms-excel.sheet.macroEnabled.12"
	default:
		return "application/octet-stream"
	}
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
