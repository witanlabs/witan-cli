package cmd

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"time"

	"github.com/coder/websocket"
)

const (
	rpcDialTimeout = 30 * time.Second
	rpcInitTimeout = 60 * time.Second
	rpcReadLimit   = 64 << 20
)

// rpcResponseEnvelope is the common envelope for RPC responses.
type rpcResponseEnvelope struct {
	ID      string          `json:"id"`
	Ok      bool            `json:"ok"`
	Type    string          `json:"type,omitempty"`
	Code    string          `json:"code,omitempty"`
	Message string          `json:"message,omitempty"`
	Meta    json.RawMessage `json:"meta,omitempty"`
}

// dialRPCWebSocket opens a WebSocket connection with the appropriate headers and auth.
// It uses the provided userAgent string and API key from the client.
func dialRPCWebSocket(ctx context.Context, wsURL string, apiKey string, userAgent string) (*websocket.Conn, error) {
	dialCtx, cancel := context.WithTimeout(ctx, rpcDialTimeout)
	defer cancel()

	headers := http.Header{}
	headers.Set("User-Agent", userAgent)

	opts := &websocket.DialOptions{HTTPHeader: headers}
	if apiKey != "" {
		opts.Subprotocols = []string{"bearer-" + apiKey}
	}

	conn, resp, err := websocket.Dial(dialCtx, wsURL, opts)
	if err != nil {
		if resp != nil {
			return nil, fmt.Errorf("opening RPC websocket: HTTP %d: %w", resp.StatusCode, err)
		}
		return nil, fmt.Errorf("opening RPC websocket: %w", err)
	}
	return conn, nil
}

// validateRPCInitResponse validates that an RPC init response indicates success.
func validateRPCInitResponse(raw []byte) error {
	var resp rpcResponseEnvelope
	if err := json.Unmarshal(raw, &resp); err != nil {
		return fmt.Errorf("parsing init response: %w", err)
	}
	if resp.Ok {
		return nil
	}
	if resp.Code != "" {
		return fmt.Errorf("%s: %s", resp.Code, resp.Message)
	}
	if resp.Message != "" {
		return fmt.Errorf("initializing RPC session: %s", resp.Message)
	}
	return fmt.Errorf("initializing RPC session failed")
}
