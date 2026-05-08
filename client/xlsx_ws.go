package client

import (
	"fmt"
	"net/url"
)

// FilesXlsxRPCWebSocketURL builds the files-backed xlsx RPC WebSocket URL.
func (c *Client) FilesXlsxRPCWebSocketURL(fileID, revisionID, hint, locale string) (string, error) {
	u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/files/"+fileID+"/xlsx/ws"))
	if err != nil {
		return "", fmt.Errorf("building URL: %w", err)
	}
	q := u.Query()
	q.Set("revision", revisionID)
	q.Set("protocol", "rpc")
	if hint != "" {
		q.Set("hint", hint)
	}
	if locale != "" {
		q.Set("locale", locale)
	}
	u.RawQuery = q.Encode()
	return httpURLToWebSocketURL(u), nil
}

// StatelessXlsxRPCWebSocketURL builds the stateless xlsx RPC WebSocket URL.
func (c *Client) StatelessXlsxRPCWebSocketURL() (string, error) {
	u, err := url.Parse(c.BaseURL + c.buildPath("v0", "/xlsx/ws"))
	if err != nil {
		return "", fmt.Errorf("building URL: %w", err)
	}
	return httpURLToWebSocketURL(u), nil
}

func httpURLToWebSocketURL(u *url.URL) string {
	switch u.Scheme {
	case "http":
		u.Scheme = "ws"
	case "https":
		u.Scheme = "wss"
	}
	return u.String()
}
