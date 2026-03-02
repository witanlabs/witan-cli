package client

import (
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"strings"
)

// detectReadContentType maps file extensions to MIME types for the read endpoint.
func detectReadContentType(filePath string) string {
	lower := strings.ToLower(filePath)
	switch {
	case strings.HasSuffix(lower, ".pdf"):
		return "application/pdf"
	case strings.HasSuffix(lower, ".docx"):
		return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
	case strings.HasSuffix(lower, ".doc"):
		return "application/msword"
	case strings.HasSuffix(lower, ".pptx"):
		return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
	case strings.HasSuffix(lower, ".ppt"):
		return "application/vnd.ms-powerpoint"
	case strings.HasSuffix(lower, ".html"), strings.HasSuffix(lower, ".htm"):
		return "text/html"
	case strings.HasSuffix(lower, ".md"):
		return "text/markdown"
	case strings.HasSuffix(lower, ".csv"):
		return "text/csv"
	case strings.HasSuffix(lower, ".tsv"):
		return "text/tab-separated-values"
	case strings.HasSuffix(lower, ".json"):
		return "application/json"
	case strings.HasSuffix(lower, ".jsonl"), strings.HasSuffix(lower, ".ndjson"):
		return "application/x-ndjson"
	case strings.HasSuffix(lower, ".xml"):
		return "application/xml"
	case strings.HasSuffix(lower, ".yaml"), strings.HasSuffix(lower, ".yml"):
		return "text/yaml"
	case strings.HasSuffix(lower, ".toml"):
		return "text/x-toml"
	default:
		return "text/plain"
	}
}

// Read calls POST /v0/read with a file in the body.
func (c *Client) Read(filePath string, params url.Values) (*ReadResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		f, err := os.Open(filePath)
		if err != nil {
			return nil, fmt.Errorf("cannot open file: %w", err)
		}

		u, err := url.Parse(c.BaseURL + "/v0/read")
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
		req.Header.Set("Content-Type", detectReadContentType(filePath))
		c.setCommonHeaders(req)
		return req, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result ReadResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing read response: %w", err)
	}
	return &result, nil
}

// ReadOutline calls POST /v0/read?outline=true with a file in the body.
func (c *Client) ReadOutline(filePath string, params url.Values) (*ReadOutlineResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		f, err := os.Open(filePath)
		if err != nil {
			return nil, fmt.Errorf("cannot open file: %w", err)
		}

		u, err := url.Parse(c.BaseURL + "/v0/read")
		if err != nil {
			f.Close()
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := make(url.Values)
		for k, v := range params {
			q[k] = v
		}
		q.Set("outline", "true")
		u.RawQuery = q.Encode()

		req, err := http.NewRequest("POST", u.String(), f)
		if err != nil {
			f.Close()
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.GetBody = func() (io.ReadCloser, error) {
			return os.Open(filePath)
		}
		req.Header.Set("Content-Type", detectReadContentType(filePath))
		c.setCommonHeaders(req)
		return req, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result ReadOutlineResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing read outline response: %w", err)
	}
	return &result, nil
}

// FilesRead calls GET /v0/files/:fileId/read.
func (c *Client) FilesRead(fileId, revisionId string, params url.Values) (*ReadResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + "/v0/files/" + fileId + "/read")
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := make(url.Values)
		for k, v := range params {
			q[k] = v
		}
		q.Set("revision", revisionId)
		u.RawQuery = q.Encode()

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

	var result ReadResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing read response: %w", err)
	}
	return &result, nil
}

// FilesReadOutline calls GET /v0/files/:fileId/read?outline=true.
func (c *Client) FilesReadOutline(fileId, revisionId string, params url.Values) (*ReadOutlineResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + "/v0/files/" + fileId + "/read")
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := make(url.Values)
		for k, v := range params {
			q[k] = v
		}
		q.Set("revision", revisionId)
		q.Set("outline", "true")
		u.RawQuery = q.Encode()

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

	var result ReadOutlineResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing read outline response: %w", err)
	}
	return &result, nil
}
