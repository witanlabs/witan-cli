package client

import (
	"context"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

type transportResult struct {
	status  int
	body    string
	headers map[string]string
	err     error
}

type sequenceTransport struct {
	t       *testing.T
	results []transportResult
	calls   int
}

func (s *sequenceTransport) RoundTrip(req *http.Request) (*http.Response, error) {
	s.calls++
	i := s.calls - 1
	if i >= len(s.results) {
		i = len(s.results) - 1
	}
	r := s.results[i]
	if r.err != nil {
		return nil, r.err
	}

	h := make(http.Header)
	for k, v := range r.headers {
		h.Set(k, v)
	}

	return &http.Response{
		StatusCode: r.status,
		Header:     h,
		Body:       io.NopCloser(strings.NewReader(r.body)),
		Request:    req,
	}, nil
}

func newTestClient(t *testing.T, tr http.RoundTripper) *Client {
	t.Helper()
	c := New("https://api.test.local", "test-key", false)
	c.HTTPClient = &http.Client{Transport: tr}
	c.sleep = func(time.Duration) {}
	c.randInt63n = func(n int64) int64 { return 0 }
	return c
}

func TestDoWithRetry_RetriesTransientStatusThenSuccess(t *testing.T) {
	tr := &sequenceTransport{
		t: t,
		results: []transportResult{
			{status: http.StatusServiceUnavailable, body: "busy"},
			{status: http.StatusBadGateway, body: "gateway"},
			{status: http.StatusOK, body: "ok"},
		},
	}
	c := newTestClient(t, tr)

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		return http.NewRequest("GET", "https://api.test.local/v0/test", nil)
	})
	if err != nil {
		t.Fatalf("doWithRetry failed: %v", err)
	}
	if tr.calls != 3 {
		t.Fatalf("expected 3 attempts, got %d", tr.calls)
	}
	if raw.StatusCode != http.StatusOK || string(raw.Body) != "ok" {
		t.Fatalf("unexpected response: status=%d body=%q", raw.StatusCode, string(raw.Body))
	}
}

func TestDoWithRetry_DoesNotRetryNonRetryableStatus(t *testing.T) {
	tr := &sequenceTransport{
		t: t,
		results: []transportResult{
			{status: http.StatusBadRequest, body: "bad"},
		},
	}
	c := newTestClient(t, tr)

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		return http.NewRequest("GET", "https://api.test.local/v0/test", nil)
	})
	if err != nil {
		t.Fatalf("doWithRetry failed: %v", err)
	}
	if tr.calls != 1 {
		t.Fatalf("expected 1 attempt, got %d", tr.calls)
	}
	if raw.StatusCode != http.StatusBadRequest {
		t.Fatalf("expected 400, got %d", raw.StatusCode)
	}
}

func TestDoWithRetry_RetriesTransportTimeoutThenSuccess(t *testing.T) {
	tr := &sequenceTransport{
		t: t,
		results: []transportResult{
			{err: &url.Error{Op: "Get", URL: "https://api.test.local/v0/test", Err: context.DeadlineExceeded}},
			{status: http.StatusOK, body: "ok"},
		},
	}
	c := newTestClient(t, tr)

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		return http.NewRequest("GET", "https://api.test.local/v0/test", nil)
	})
	if err != nil {
		t.Fatalf("doWithRetry failed: %v", err)
	}
	if tr.calls != 2 {
		t.Fatalf("expected 2 attempts, got %d", tr.calls)
	}
	if raw.StatusCode != http.StatusOK {
		t.Fatalf("expected 200, got %d", raw.StatusCode)
	}
}

func TestDoWithRetry_HonorsRetryAfterHeader(t *testing.T) {
	tr := &sequenceTransport{
		t: t,
		results: []transportResult{
			{status: http.StatusTooManyRequests, body: "rate limited", headers: map[string]string{"Retry-After": "2"}},
			{status: http.StatusOK, body: "ok"},
		},
	}
	c := newTestClient(t, tr)

	var slept []time.Duration
	c.sleep = func(d time.Duration) {
		slept = append(slept, d)
	}

	_, err := c.doWithRetry(func() (*http.Request, error) {
		return http.NewRequest("GET", "https://api.test.local/v0/test", nil)
	})
	if err != nil {
		t.Fatalf("doWithRetry failed: %v", err)
	}
	if len(slept) != 1 {
		t.Fatalf("expected one sleep, got %d", len(slept))
	}
	if slept[0] != 2*time.Second {
		t.Fatalf("expected sleep of 2s, got %s", slept[0])
	}
}

func TestDoWithRetry_ReturnsRetryAfterOnTerminalRateLimit(t *testing.T) {
	tr := &sequenceTransport{
		t: t,
		results: []transportResult{
			{status: http.StatusTooManyRequests, body: "rate limited", headers: map[string]string{"Retry-After": "7"}},
		},
	}
	c := newTestClient(t, tr)
	c.maxAttempts = 1

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		return http.NewRequest("GET", "https://api.test.local/v0/test", nil)
	})
	if err != nil {
		t.Fatalf("doWithRetry failed: %v", err)
	}
	if raw.StatusCode != http.StatusTooManyRequests {
		t.Fatalf("expected 429, got %d", raw.StatusCode)
	}
	if raw.RetryAfter != "7" {
		t.Fatalf("expected Retry-After header to be preserved, got %q", raw.RetryAfter)
	}
}

func TestParseAPIError_RateLimitMessage(t *testing.T) {
	err := parseAPIError(http.StatusTooManyRequests, []byte(`{"error":{"message":"too many requests","code":"rate_limited"}}`), "9")
	apiErr, ok := err.(*APIError)
	if !ok {
		t.Fatalf("expected APIError, got %T", err)
	}
	if got := apiErr.Error(); got != "rate limited by API; retry after 9" {
		t.Fatalf("unexpected rate-limit message: %q", got)
	}

	err = parseAPIError(http.StatusTooManyRequests, []byte("rate limited"), "")
	apiErr, ok = err.(*APIError)
	if !ok {
		t.Fatalf("expected APIError, got %T", err)
	}
	if got := apiErr.Error(); got != "rate limited by API; retry in a moment" {
		t.Fatalf("unexpected rate-limit fallback message: %q", got)
	}
}

func TestUploadFile_RetriesAndReplaysMultipartBody(t *testing.T) {
	tr := &sequenceTransport{
		t: t,
		results: []transportResult{
			{status: http.StatusServiceUnavailable, body: "try again"},
			{
				status: http.StatusOK,
				body:   `{"id":"file_1","object":"file","filename":"test.xlsx","bytes":3,"revision_id":"rev_1","status":"processed"}`,
			},
		},
	}

	var bodies []string
	clientTransport := roundTripperFunc(func(req *http.Request) (*http.Response, error) {
		if req.Method != "POST" {
			t.Fatalf("expected POST, got %s", req.Method)
		}
		b, err := io.ReadAll(req.Body)
		if err != nil {
			t.Fatalf("reading request body: %v", err)
		}
		bodies = append(bodies, string(b))
		return tr.RoundTrip(req)
	})

	c := newTestClient(t, clientTransport)

	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "test.xlsx")
	if err := os.WriteFile(filePath, []byte("abc"), 0o644); err != nil {
		t.Fatalf("writing temp file: %v", err)
	}

	resp, err := c.UploadFile(filePath)
	if err != nil {
		t.Fatalf("UploadFile failed: %v", err)
	}
	if tr.calls != 2 {
		t.Fatalf("expected 2 attempts, got %d", tr.calls)
	}
	if resp.ID != "file_1" || resp.RevisionID != "rev_1" {
		t.Fatalf("unexpected upload response: %+v", resp)
	}
	if len(bodies) != 2 {
		t.Fatalf("expected 2 request bodies, got %d", len(bodies))
	}
	for i, body := range bodies {
		if !strings.Contains(body, `name="file"; filename="test.xlsx"`) {
			t.Fatalf("request %d missing multipart filename", i+1)
		}
		if !strings.Contains(body, "abc") {
			t.Fatalf("request %d missing file content", i+1)
		}
	}
}

type roundTripperFunc func(*http.Request) (*http.Response, error)

func (f roundTripperFunc) RoundTrip(req *http.Request) (*http.Response, error) {
	return f(req)
}
