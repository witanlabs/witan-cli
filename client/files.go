package client

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"mime"
	"mime/multipart"
	"net/http"
	"net/textproto"
	"net/url"
	"os"
	"path/filepath"
)

// FileResponse is the response from POST /v0/files.
type FileResponse struct {
	ID         string `json:"id"`
	Object     string `json:"object"`
	Filename   string `json:"filename"`
	Bytes      int64  `json:"bytes"`
	RevisionID string `json:"revision_id"`
	Status     string `json:"status"`
}

// UploadFile uploads a local file via multipart POST to /v0/files
// and returns the file metadata including fileId and revisionId.
func (c *Client) UploadFile(filePath string) (*FileResponse, error) {
	payload, contentType, err := buildMultipartPayload(filePath)
	if err != nil {
		return nil, err
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		req, err := http.NewRequest("POST", c.BaseURL+"/v0/files", bytes.NewReader(payload))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.Header.Set("Content-Type", contentType)
		setAuthorization(req, c.APIKey)
		return req, nil
	})
	if err != nil {
		return nil, err
	}

	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result FileResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing upload response: %w", err)
	}
	return &result, nil
}

// UploadFileVersion uploads a local file as a new revision of an existing file.
func (c *Client) UploadFileVersion(fileID, filePath string) (*FileResponse, error) {
	payload, contentType, err := buildMultipartPayload(filePath)
	if err != nil {
		return nil, err
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		req, err := http.NewRequest("PUT", c.BaseURL+"/v0/files/"+fileID, bytes.NewReader(payload))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.Header.Set("Content-Type", contentType)
		setAuthorization(req, c.APIKey)
		return req, nil
	})
	if err != nil {
		return nil, err
	}

	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result FileResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing upload response: %w", err)
	}
	return &result, nil
}

func buildMultipartPayload(filePath string) ([]byte, string, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, "", fmt.Errorf("cannot open file: %w", err)
	}
	defer f.Close()

	var buf bytes.Buffer
	writer := multipart.NewWriter(&buf)

	filename := filepath.Base(filePath)
	mimeType := mime.TypeByExtension(filepath.Ext(filePath))
	if mimeType == "" {
		mimeType = "application/octet-stream"
	}
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
	if err := writer.Close(); err != nil {
		return nil, "", fmt.Errorf("finalizing multipart payload: %w", err)
	}

	return buf.Bytes(), writer.FormDataContentType(), nil
}

// EnsureUploaded hashes the file, checks the cache, uploads if needed,
// and returns (fileId, revisionId). On a cache hit that turns out to be a 404,
// it evicts and re-uploads.
func (c *Client) EnsureUploaded(filePath string) (fileId, revisionId string, err error) {
	if c.cache == nil {
		// No cache (stateless) — upload every time
		resp, err := c.UploadFile(filePath)
		if err != nil {
			return "", "", err
		}
		return resp.ID, resp.RevisionID, nil
	}

	key, err := HashFile(filePath, c.BaseURL)
	if err != nil {
		return "", "", err
	}

	if entry, ok := c.cache.Get(key); ok {
		c.cache.PutKnown(filePath, c.BaseURL, entry)
		return entry.FileID, entry.RevisionID, nil
	}

	// Cache miss. If this local file is known, upload as a new revision first.
	if known, ok := c.cache.GetKnown(filePath, c.BaseURL); ok {
		resp, err := c.UploadFileVersion(known.FileID, filePath)
		if err == nil {
			entry := cacheEntryFromUpload(resp)
			c.cache.Put(key, entry)
			c.cache.PutKnown(filePath, c.BaseURL, entry)
			return resp.ID, resp.RevisionID, nil
		}
		if !shouldFallbackToFreshUpload(err) {
			return "", "", err
		}
	}

	// No known file (or version upload rejected) — create a new file.
	resp, err := c.UploadFile(filePath)
	if err != nil {
		return "", "", err
	}

	entry := cacheEntryFromUpload(resp)
	c.cache.Put(key, entry)
	c.cache.PutKnown(filePath, c.BaseURL, entry)
	return resp.ID, resp.RevisionID, nil
}

// ReuploadFile evicts the cache entry for the given file and re-uploads it.
// Use this after getting a 404 from a files endpoint (stale cache entry).
func (c *Client) ReuploadFile(filePath string) (fileId, revisionId string, err error) {
	if c.cache != nil {
		key, err := HashFile(filePath, c.BaseURL)
		if err == nil {
			c.cache.Evict(key)
		}
		c.cache.EvictKnown(filePath, c.BaseURL)
	}
	return c.EnsureUploaded(filePath)
}

// UpdateCachedRevision updates hash and local-file cache entries after a command
// produces a new revision for an already-known file.
func (c *Client) UpdateCachedRevision(filePath, fileID, revisionID string) error {
	if c.cache == nil {
		return nil
	}

	key, err := HashFile(filePath, c.BaseURL)
	if err != nil {
		return err
	}

	entry := CacheEntry{
		FileID:     fileID,
		RevisionID: revisionID,
		Filename:   filepath.Base(filePath),
	}
	if fi, statErr := os.Stat(filePath); statErr == nil {
		entry.Bytes = fi.Size()
	}

	c.cache.Put(key, entry)
	c.cache.PutKnown(filePath, c.BaseURL, entry)
	return nil
}

func cacheEntryFromUpload(resp *FileResponse) CacheEntry {
	return CacheEntry{
		FileID:     resp.ID,
		RevisionID: resp.RevisionID,
		Bytes:      resp.Bytes,
		Filename:   resp.Filename,
	}
}

func shouldFallbackToFreshUpload(err error) bool {
	apiErr, ok := err.(*APIError)
	if !ok {
		return false
	}
	if apiErr.StatusCode == 404 {
		return true
	}
	return apiErr.Code == "filename_mismatch" || apiErr.Code == "content_type_mismatch"
}

// FilesLint calls GET /v0/files/:fileId/xlsx/lint and returns lint diagnostics.
func (c *Client) FilesLint(fileId, revisionId string, params url.Values) (*LintResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + "/v0/files/" + fileId + "/xlsx/lint")
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
		setAuthorization(req, c.APIKey)
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

// FilesCalc calls GET /v0/files/:fileId/xlsx/calc and returns calc results.
func (c *Client) FilesCalc(fileId, revisionId string, params url.Values) (*CalcResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + "/v0/files/" + fileId + "/xlsx/calc")
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
		setAuthorization(req, c.APIKey)
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

// FilesEdit calls POST /v0/files/:fileId/xlsx/edit with JSON body and returns edit results.
func (c *Client) FilesEdit(fileId, revisionId string, cells []EditCell) (*EditResponse, error) {
	body, err := json.Marshal(map[string]any{"cells": cells})
	if err != nil {
		return nil, fmt.Errorf("marshaling edit body: %w", err)
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + "/v0/files/" + fileId + "/xlsx/edit")
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := u.Query()
		q.Set("revision", revisionId)
		u.RawQuery = q.Encode()

		req, err := http.NewRequest("POST", u.String(), bytes.NewReader(body))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.Header.Set("Content-Type", "application/json")
		setAuthorization(req, c.APIKey)
		return req, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result EditResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing edit response: %w", err)
	}
	return &result, nil
}

// DownloadFileContent calls GET /v0/files/:fileId/content and returns the raw file bytes.
func (c *Client) DownloadFileContent(fileId, revisionId string) ([]byte, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + "/v0/files/" + fileId + "/content")
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		if revisionId != "" {
			q := u.Query()
			q.Set("revision", revisionId)
			u.RawQuery = q.Encode()
		}

		req, err := http.NewRequest("GET", u.String(), nil)
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		setAuthorization(req, c.APIKey)
		return req, nil
	})
	if err != nil {
		return nil, err
	}
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}
	return raw.Body, nil
}

// FilesRender calls GET /v0/files/:fileId/xlsx/render and returns image bytes.
func (c *Client) FilesRender(fileId, revisionId string, params map[string]string) ([]byte, string, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + "/v0/files/" + fileId + "/xlsx/render")
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := u.Query()
		q.Set("revision", revisionId)
		for k, v := range params {
			q.Set(k, v)
		}
		u.RawQuery = q.Encode()

		req, err := http.NewRequest("GET", u.String(), nil)
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		setAuthorization(req, c.APIKey)
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
