package client

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"net/textproto"
	"net/url"
	"os"
	"path/filepath"
	"strings"
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
		req, err := http.NewRequest("POST", c.BaseURL+c.buildPath("v0", "/files"), bytes.NewReader(payload))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.Header.Set("Content-Type", contentType)
		c.setCommonHeaders(req)
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
		req, err := http.NewRequest("PUT", c.BaseURL+c.buildPath("v0", "/files/"+fileID), bytes.NewReader(payload))
		if err != nil {
			return nil, fmt.Errorf("creating request: %w", err)
		}
		req.Header.Set("Content-Type", contentType)
		c.setCommonHeaders(req)
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
	if err := writer.Close(); err != nil {
		return nil, "", fmt.Errorf("finalizing multipart payload: %w", err)
	}

	return buf.Bytes(), writer.FormDataContentType(), nil
}

// EnsureUploaded looks up the file by path in the cache and returns
// (fileId, revisionId). If the cached entry's content hash matches the
// current file, the cached pair is returned. If the file has changed,
// a new revision is PUT under the same fileID; if that PUT fails because
// the fileID is gone (or the server rejects the version), it falls back
// to a fresh POST. With no cache entry, a fresh POST is made.
//
// On a 404 from a downstream op, the caller should call ReuploadFile,
// which evicts and runs through this path again.
func (c *Client) EnsureUploaded(filePath string) (fileId, revisionId string, err error) {
	if c.cache == nil {
		// No cache (stateless) — upload every time
		resp, err := c.UploadFile(filePath)
		if err != nil {
			return "", "", err
		}
		return resp.ID, resp.RevisionID, nil
	}

	if entry, ok := c.cache.Get(filePath, c.BaseURL, c.OrgID); ok {
		hash, err := hashFile(filePath)
		if err != nil {
			return "", "", err
		}
		if hash == entry.ContentHash {
			return entry.FileID, entry.RevisionID, nil
		}

		resp, err := c.UploadFileVersion(entry.FileID, filePath)
		if err == nil {
			c.cache.Put(filePath, c.BaseURL, c.OrgID, cacheEntryFromUpload(resp, hash))
			return resp.ID, resp.RevisionID, nil
		}
		if !shouldFallbackToFreshUpload(err) {
			return "", "", err
		}
		// Fall through to fresh POST.
	}

	resp, err := c.UploadFile(filePath)
	if err != nil {
		return "", "", err
	}

	hash, err := hashFile(filePath)
	if err != nil {
		return "", "", err
	}
	c.cache.Put(filePath, c.BaseURL, c.OrgID, cacheEntryFromUpload(resp, hash))
	return resp.ID, resp.RevisionID, nil
}

// ReuploadFile evicts the cache entry for the given file and re-uploads it.
// Use this after getting a 404 from a files endpoint (stale cache entry).
func (c *Client) ReuploadFile(filePath string) (fileId, revisionId string, err error) {
	if c.cache != nil {
		c.cache.Evict(filePath, c.BaseURL, c.OrgID)
	}
	return c.EnsureUploaded(filePath)
}

// UpdateCachedRevision updates the cache entry after a command produces a new
// revision for the given file path.
func (c *Client) UpdateCachedRevision(filePath, fileID, revisionID string) error {
	if c.cache == nil {
		return nil
	}

	hash, err := hashFile(filePath)
	if err != nil {
		return err
	}

	entry := CacheEntry{
		FileID:      fileID,
		RevisionID:  revisionID,
		ContentHash: hash,
		Filename:    filepath.Base(filePath),
	}
	if fi, statErr := os.Stat(filePath); statErr == nil {
		entry.Bytes = fi.Size()
	}

	c.cache.Put(filePath, c.BaseURL, c.OrgID, entry)
	return nil
}

func cacheEntryFromUpload(resp *FileResponse, contentHash string) CacheEntry {
	return CacheEntry{
		FileID:      resp.ID,
		RevisionID:  resp.RevisionID,
		ContentHash: contentHash,
		Bytes:       resp.Bytes,
		Filename:    resp.Filename,
	}
}

// knownMIMEType returns the MIME type for file extensions we handle, without
// consulting the system MIME database. Returns "" for unknown extensions.
func knownMIMEType(ext string) string {
	switch strings.ToLower(ext) {
	case ".xlsx":
		return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	case ".xls":
		return "application/vnd.ms-excel"
	case ".xlsm":
		return "application/vnd.ms-excel.sheet.macroEnabled.12"
	case ".csv":
		return "text/csv"
	case ".pdf":
		return "application/pdf"
	case ".docx":
		return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
	case ".doc":
		return "application/msword"
	case ".pptx":
		return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
	case ".ppt":
		return "application/vnd.ms-powerpoint"
	case ".json":
		return "application/json"
	case ".xml":
		return "application/xml"
	case ".html", ".htm":
		return "text/html"
	default:
		return ""
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
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileId+"/xlsx/lint"))
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

	var result LintResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing lint response: %w", err)
	}
	return &result, nil
}

// FilesCalc calls GET /v0/files/:fileId/xlsx/calc and returns calc results.
func (c *Client) FilesCalc(fileId, revisionId string, params url.Values) (*CalcResponse, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileId+"/xlsx/calc"))
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

	var result CalcResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing calc response: %w", err)
	}
	return &result, nil
}

// FilesExec calls POST /v0/files/:fileId/xlsx/exec with JSON body and returns exec results.
func (c *Client) FilesExec(fileID, revisionID string, req ExecRequest, save bool) (*ExecResponse, error) {
	body, err := json.Marshal(req)
	if err != nil {
		return nil, fmt.Errorf("marshaling exec body: %w", err)
	}

	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileID+"/xlsx/exec"))
		if err != nil {
			return nil, fmt.Errorf("building URL: %w", err)
		}
		q := u.Query()
		q.Set("revision", revisionID)
		q.Set("cache", "true")
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
	if raw.StatusCode != 200 {
		return nil, parseAPIError(raw.StatusCode, raw.Body, raw.RetryAfter)
	}

	var result ExecResponse
	if err := json.Unmarshal(raw.Body, &result); err != nil {
		return nil, fmt.Errorf("parsing exec response: %w", err)
	}
	return &result, nil
}

// DownloadFileContent calls GET /v0/files/:fileId/content and returns the raw file bytes.
func (c *Client) DownloadFileContent(fileId, revisionId string) ([]byte, error) {
	raw, err := c.doWithRetry(func() (*http.Request, error) {
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileId+"/content"))
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
		c.setCommonHeaders(req)
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
		u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileId+"/xlsx/render"))
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
