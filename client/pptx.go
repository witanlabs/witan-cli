package client

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
)

// PPTXRender renders a PPTX slide and returns the image bytes.
func (c *Client) PPTXRender(filePath string, params map[string]string) ([]byte, string, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		f, err := os.Open(filePath)
		if err != nil {
			return nil, fmt.Errorf("cannot open file: %w", err)
		}

		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/pptx/render"))
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
	if raw.StatusCode != http.StatusOK {
		return nil, "", parsePPTXAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}
	return raw.Body, raw.ContentType, nil
}

// PPTXExec runs Office.js-compatible JavaScript against a PPTX file via
// multipart POST /v0/pptx/exec.
func (c *Client) PPTXExec(filePath string, req ExecRequest, save bool) (*ExecResponse, error) {
	payload, contentType, err := buildExecMultipartPayload(filePath, req, true)
	if err != nil {
		return nil, err
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/pptx/exec"))
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
	if raw.StatusCode != http.StatusOK {
		return nil, parsePPTXAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result ExecResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing exec response: %w", err)
	}
	return &result, nil
}

// PPTXExecCreate runs Office.js-compatible JavaScript against a new empty PPTX
// file via multipart POST /v0/pptx/exec?create=true.
func (c *Client) PPTXExecCreate(filePath string, req ExecRequest, save bool) (*ExecResponse, error) {
	if req.Filename == "" {
		req.Filename = filepath.Base(filePath)
	}
	payload, contentType, err := buildExecMultipartPayload(filePath, req, false)
	if err != nil {
		return nil, err
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/pptx/exec"))
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
	if raw.StatusCode != http.StatusOK {
		return nil, parsePPTXAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result ExecResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing exec response: %w", err)
	}
	return &result, nil
}

// FilesPPTXExec calls POST /v0/files/:fileId/pptx/exec with a JSON body.
func (c *Client) FilesPPTXExec(fileID, revisionID string, req ExecRequest, save bool) (*ExecResponse, error) {
	body, err := json.Marshal(req)
	if err != nil {
		return nil, fmt.Errorf("marshaling exec body: %w", err)
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileID+"/pptx/exec"))
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := u.Query()
		q.Set("revision", revisionID)
		if save {
			q.Set("save", "true")
		}
		if req.Locale != "" {
			q.Set("locale", req.Locale)
		}
		u.RawQuery = q.Encode()

		httpReq, err := http.NewRequest("POST", u.String(), bytes.NewReader(body))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		httpReq.Header.Set("Content-Type", "application/json")
		c.setCommonHeaders(httpReq)
		if req.Locale != "" {
			httpReq.Header.Set("Accept-Language", req.Locale)
		}
		return httpReq, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != http.StatusOK {
		return nil, parsePPTXAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result ExecResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing exec response: %w", err)
	}
	return &result, nil
}

// FilesPPTXRender calls GET /v0/files/:fileId/pptx/render and returns image bytes.
func (c *Client) FilesPPTXRender(fileID, revisionID string, params map[string]string) ([]byte, string, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileID+"/pptx/render"))
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := u.Query()
		q.Set("revision", revisionID)
		for k, v := range params {
			q.Set(k, v)
		}
		u.RawQuery = q.Encode()

		req, err := http.NewRequest("GET", u.String(), nil)
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		c.setCommonHeaders(req)
		return req, nil
	})
	if err != nil {
		return nil, "", err
	}
	if raw.StatusCode != http.StatusOK {
		return nil, "", parsePPTXAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}
	return raw.Body, raw.ContentType, nil
}

// PPTXExecTypes fetches the combined TypeScript declarations for the pptx exec
// sandbox (stripped Office.js surface plus Witan chart extensions) via GET
// /v0/pptx/exec/types. The endpoint is public and returns raw text/plain; no
// auth headers are required.
func (c *Client) PPTXExecTypes() ([]byte, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/pptx/exec/types"))
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
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
	if raw.StatusCode != http.StatusOK {
		return nil, parsePPTXAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}
	return raw.Body, nil
}

// PPTXLint lints a PPTX file via POST /v0/pptx/lint.
func (c *Client) PPTXLint(filePath string, params url.Values) (*PptxLintResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		f, err := os.Open(filePath)
		if err != nil {
			return nil, fmt.Errorf("cannot open file: %w", err)
		}

		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/pptx/lint"))
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
	if raw.StatusCode != http.StatusOK {
		return nil, parsePPTXAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result PptxLintResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing lint response: %w", err)
	}
	return &result, nil
}

// FilesPPTXLint calls GET /v0/files/:fileId/pptx/lint.
func (c *Client) FilesPPTXLint(fileID, revisionID string, params url.Values) (*PptxLintResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileID+"/pptx/lint"))
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := make(url.Values)
		for k, v := range params {
			q[k] = v
		}
		q.Set("revision", revisionID)
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
	if raw.StatusCode != http.StatusOK {
		return nil, parsePPTXAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result PptxLintResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing lint response: %w", err)
	}
	return &result, nil
}

func parsePPTXAPIError(statusCode int, body []byte, retryAfter string) error {
	err := parseAPIError(statusCode, body, retryAfter)
	apiErr, ok := err.(*APIError)
	if !ok || apiErr.Code != "invalid_mime_type" {
		return err
	}
	return &APIError{
		StatusCode: apiErr.StatusCode,
		Code:       apiErr.Code,
		Message:    "unsupported file type - expected .pptx",
		RetryAfter: apiErr.RetryAfter,
	}
}
