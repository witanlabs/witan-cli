package cmd

import (
	"bufio"
	"context"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"strings"
	"time"

	"github.com/coder/websocket"
	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	rpcHint   string
	rpcLocale string
)

const (
	rpcDialTimeout = 30 * time.Second
	rpcInitTimeout = 60 * time.Second
	rpcReadLimit   = 64 << 20
)

var xlsxRPCCmd = &cobra.Command{
	Use:   "rpc <file>",
	Short: "Run newline-delimited xlsx RPC over stdio",
	Long: `Run newline-delimited xlsx RPC over stdio.

The command opens <file> as a WebSocket-backed workbook session, then relays
one JSON object per input line to the xlsx RPC endpoint. Stdout contains one
redacted JSON response per request. Stderr is reserved for CLI diagnostics.

Input shape:
  {"id":"1","op":"listSheets","args":{}}
  {"id":"2","op":"readRange","args":{"address":"Sheet1!A1:B10"}}
  {"id":"3","op":"save","args":{}}

The CLI owns session setup. Do not include a workbook field. Save metadata
returned by the API is used for local writeback and omitted from stdout.`,
	Args: cobra.ExactArgs(1),
	RunE: runRPC,
}

type rpcRequestEnvelope struct {
	ID string `json:"id"`
	Op string `json:"op"`
}

type rpcResponseEnvelope struct {
	ID      string          `json:"id"`
	Ok      bool            `json:"ok"`
	Code    string          `json:"code"`
	Message string          `json:"message"`
	Meta    json.RawMessage `json:"meta"`
}

type rpcResponseMeta struct {
	RevisionID  string `json:"revision_id"`
	File        string `json:"file"`
	ContentType string `json:"content_type"`
}

type statelessRPCInitMessage struct {
	Type        string `json:"type"`
	ID          string `json:"id"`
	ContentType string `json:"content_type,omitempty"`
	File        string `json:"file,omitempty"`
	Hint        string `json:"hint,omitempty"`
	Locale      string `json:"locale,omitempty"`
}

type rpcSession struct {
	conn     *websocket.Conn
	client   *client.Client
	filePath string
	fileID   string
	hint     string
	locale   string
	mode     string
}

func init() {
	xlsxRPCCmd.Flags().StringVar(&rpcHint, "hint", "", "Sheet name or address hint for lazy workbook loading")
	xlsxRPCCmd.Flags().StringVar(&rpcLocale, "locale", "", "Execution locale (env: WITAN_LOCALE; otherwise LC_ALL / LC_MESSAGES / LANG)")
	xlsxCmd.AddCommand(xlsxRPCCmd)
}

func runRPC(cmd *cobra.Command, args []string) error {
	cmd.SilenceUsage = true
	filePath := args[0]

	filePath, err := fixExcelExtension(filePath)
	if err != nil {
		return err
	}

	locale, err := resolveRPCLocale(cmd)
	if err != nil {
		return err
	}

	key, orgID, err := resolveAuth()
	if err != nil {
		return err
	}
	c := newAPIClient(key, orgID)

	session, err := openRPCSession(cmd.Context(), c, filePath, rpcHint, locale)
	if err != nil {
		return err
	}
	defer session.close()

	return relayRPCStdio(cmd.Context(), session, os.Stdin, os.Stdout)
}

func openRPCSession(ctx context.Context, c *client.Client, filePath, hint, locale string) (*rpcSession, error) {
	if c.Stateless {
		return openStatelessRPCSession(ctx, c, filePath, hint, locale)
	}
	return openFilesRPCSession(ctx, c, filePath, hint, locale)
}

func openFilesRPCSession(ctx context.Context, c *client.Client, filePath, hint, locale string) (*rpcSession, error) {
	fileID, revisionID, err := c.EnsureUploaded(filePath)
	if err != nil {
		return nil, err
	}

	wsURL, err := c.FilesXlsxRPCWebSocketURL(fileID, revisionID, hint, locale)
	if err != nil {
		return nil, err
	}
	conn, err := dialRPCWebSocket(ctx, c, wsURL)
	if err != nil {
		return nil, err
	}
	conn.SetReadLimit(rpcReadLimit)
	return &rpcSession{
		conn:     conn,
		client:   c,
		filePath: filePath,
		fileID:   fileID,
		hint:     hint,
		locale:   locale,
		mode:     "files",
	}, nil
}

func openStatelessRPCSession(ctx context.Context, c *client.Client, filePath, hint, locale string) (*rpcSession, error) {
	wsURL, err := c.StatelessXlsxRPCWebSocketURL()
	if err != nil {
		return nil, err
	}
	conn, err := dialRPCWebSocket(ctx, c, wsURL)
	if err != nil {
		return nil, err
	}
	conn.SetReadLimit(rpcReadLimit)

	b, err := os.ReadFile(filePath)
	if err != nil {
		conn.CloseNow()
		return nil, fmt.Errorf("reading workbook: %w", err)
	}
	initMsg := statelessRPCInitMessage{
		Type:        "init",
		ID:          "witan-init-1",
		ContentType: client.DetectContentType(filePath),
		File:        base64.StdEncoding.EncodeToString(b),
		Hint:        hint,
		Locale:      locale,
	}
	raw, err := json.Marshal(initMsg)
	if err != nil {
		conn.CloseNow()
		return nil, fmt.Errorf("marshaling init message: %w", err)
	}

	initCtx, cancel := context.WithTimeout(ctx, rpcInitTimeout)
	defer cancel()
	if err := conn.Write(initCtx, websocket.MessageText, raw); err != nil {
		conn.CloseNow()
		return nil, fmt.Errorf("sending init message: %w", err)
	}
	msgType, resp, err := conn.Read(initCtx)
	if err != nil {
		conn.CloseNow()
		return nil, fmt.Errorf("reading init response: %w", err)
	}
	if msgType != websocket.MessageText {
		conn.CloseNow()
		return nil, fmt.Errorf("reading init response: expected text frame, got %v", msgType)
	}
	if err := validateRPCInitResponse(resp); err != nil {
		conn.CloseNow()
		return nil, err
	}

	return &rpcSession{
		conn:     conn,
		client:   c,
		filePath: filePath,
		mode:     "stateless",
	}, nil
}

func dialRPCWebSocket(ctx context.Context, c *client.Client, wsURL string) (*websocket.Conn, error) {
	dialCtx, cancel := context.WithTimeout(ctx, rpcDialTimeout)
	defer cancel()

	headers := http.Header{}
	headers.Set("User-Agent", cliUserAgent())

	opts := &websocket.DialOptions{HTTPHeader: headers}
	if c.APIKey != "" {
		opts.Subprotocols = []string{"bearer-" + c.APIKey}
	}

	conn, resp, err := websocket.Dial(dialCtx, wsURL, opts)
	if err != nil {
		if resp != nil {
			return nil, fmt.Errorf("opening xlsx RPC websocket: HTTP %d: %w", resp.StatusCode, err)
		}
		return nil, fmt.Errorf("opening xlsx RPC websocket: %w", err)
	}
	return conn, nil
}

func (s *rpcSession) close() {
	if s.conn != nil {
		_ = s.conn.Close(websocket.StatusNormalClosure, "")
	}
}

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
		return fmt.Errorf("initializing xlsx RPC session: %s", resp.Message)
	}
	return fmt.Errorf("initializing xlsx RPC session failed")
}

func relayRPCStdio(ctx context.Context, session *rpcSession, stdin io.Reader, stdout io.Writer) error {
	scanner := bufio.NewScanner(stdin)
	scanner.Buffer(make([]byte, 0, 64*1024), rpcReadLimit)

	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "" {
			continue
		}

		var req rpcRequestEnvelope
		_ = json.Unmarshal([]byte(line), &req)

		redacted, err := session.sendRPCLine(ctx, req, []byte(line))
		if err != nil {
			return err
		}
		if _, err := stdout.Write(redacted); err != nil {
			return fmt.Errorf("writing RPC response: %w", err)
		}
		if _, err := stdout.Write([]byte("\n")); err != nil {
			return fmt.Errorf("writing RPC response: %w", err)
		}
	}
	if err := scanner.Err(); err != nil {
		return fmt.Errorf("reading RPC stdin: %w", err)
	}
	return nil
}

func (s *rpcSession) sendRPCLine(ctx context.Context, req rpcRequestEnvelope, line []byte) ([]byte, error) {
	for attempt := 0; attempt < 2; attempt++ {
		msgType, rawResp, err := s.roundTripRPCLine(ctx, line)
		if err != nil {
			if attempt == 0 && s.isFilesStaleCacheReadError(err) {
				if reconnectErr := s.reconnectFilesRPCSession(ctx); reconnectErr != nil {
					return nil, reconnectErr
				}
				continue
			}
			return nil, err
		}
		if msgType != websocket.MessageText {
			return nil, fmt.Errorf("reading RPC response: expected text frame, got %v", msgType)
		}
		if attempt == 0 && s.isFilesStaleCacheResponse(req, rawResp) {
			if err := s.reconnectFilesRPCSession(ctx); err != nil {
				return nil, err
			}
			continue
		}
		return s.applyRPCResponseSideEffects(req, rawResp)
	}
	return nil, fmt.Errorf("reconnecting stale xlsx RPC session failed")
}

func (s *rpcSession) roundTripRPCLine(ctx context.Context, line []byte) (websocket.MessageType, []byte, error) {
	if err := s.conn.Write(ctx, websocket.MessageText, line); err != nil {
		return 0, nil, fmt.Errorf("sending RPC message: %w", err)
	}

	msgType, rawResp, err := s.conn.Read(ctx)
	if err != nil {
		return 0, nil, fmt.Errorf("reading RPC response: %w", err)
	}
	return msgType, rawResp, nil
}

func (s *rpcSession) reconnectFilesRPCSession(ctx context.Context) error {
	if s.mode != "files" {
		return fmt.Errorf("cannot reconnect %s RPC session", s.mode)
	}
	if s.conn != nil {
		s.conn.CloseNow()
	}

	fileID, revisionID, err := s.client.ReuploadFile(s.filePath)
	if err != nil {
		return fmt.Errorf("reuploading workbook after stale RPC session: %w", err)
	}
	wsURL, err := s.client.FilesXlsxRPCWebSocketURL(fileID, revisionID, s.hint, s.locale)
	if err != nil {
		return err
	}
	conn, err := dialRPCWebSocket(ctx, s.client, wsURL)
	if err != nil {
		return err
	}
	conn.SetReadLimit(rpcReadLimit)
	s.conn = conn
	s.fileID = fileID
	return nil
}

func (s *rpcSession) isFilesStaleCacheResponse(req rpcRequestEnvelope, rawResp []byte) bool {
	if s.mode != "files" {
		return false
	}
	var resp rpcResponseEnvelope
	if err := json.Unmarshal(rawResp, &resp); err != nil {
		return false
	}
	if resp.ID != "" {
		return false
	}
	return isFilesStaleCacheCode(resp.Code)
}

func (s *rpcSession) isFilesStaleCacheReadError(err error) bool {
	return s.mode == "files" && websocket.CloseStatus(err) == 4404
}

func isFilesStaleCacheCode(code string) bool {
	switch strings.ToUpper(code) {
	case "FILE_NOT_FOUND", "REVISION_NOT_FOUND":
		return true
	default:
		return false
	}
}

func (s *rpcSession) applyRPCResponseSideEffects(req rpcRequestEnvelope, rawResp []byte) ([]byte, error) {
	var resp rpcResponseEnvelope
	if err := json.Unmarshal(rawResp, &resp); err != nil {
		return nil, fmt.Errorf("parsing RPC response: %w", err)
	}

	if resp.Ok && strings.EqualFold(req.Op, "save") {
		if err := s.applySaveResponse(resp); err != nil {
			return nil, err
		}
	}

	var obj map[string]json.RawMessage
	if err := json.Unmarshal(rawResp, &obj); err != nil {
		return nil, fmt.Errorf("parsing RPC response: %w", err)
	}
	delete(obj, "meta")
	redacted, err := json.Marshal(obj)
	if err != nil {
		return nil, fmt.Errorf("redacting RPC response: %w", err)
	}
	return redacted, nil
}

func (s *rpcSession) applySaveResponse(resp rpcResponseEnvelope) error {
	if len(resp.Meta) == 0 {
		return fmt.Errorf("save response missing transport metadata")
	}
	var meta rpcResponseMeta
	if err := json.Unmarshal(resp.Meta, &meta); err != nil {
		return fmt.Errorf("parsing save metadata: %w", err)
	}

	switch s.mode {
	case "files":
		if meta.RevisionID == "" {
			return fmt.Errorf("save response missing revision_id metadata")
		}
		fileBytes, err := s.client.DownloadFileContent(s.fileID, meta.RevisionID)
		if err != nil {
			return fmt.Errorf("downloading saved workbook: %w", err)
		}
		if err := os.WriteFile(s.filePath, fileBytes, 0o644); err != nil {
			return fmt.Errorf("writing saved workbook: %w", err)
		}
		newPath, err := fixWritebackExtension(s.filePath)
		if err != nil {
			return err
		}
		s.filePath = newPath
		if err := s.client.UpdateCachedRevision(s.filePath, s.fileID, meta.RevisionID); err != nil {
			return fmt.Errorf("updating local cache: %w", err)
		}
	case "stateless":
		if meta.File == "" {
			return fmt.Errorf("save response missing file metadata")
		}
		decoded, err := base64.StdEncoding.DecodeString(meta.File)
		if err != nil {
			return fmt.Errorf("decoding saved workbook: %w", err)
		}
		if err := os.WriteFile(s.filePath, decoded, 0o644); err != nil {
			return fmt.Errorf("writing saved workbook: %w", err)
		}
		newPath, err := fixWritebackExtension(s.filePath)
		if err != nil {
			return err
		}
		s.filePath = newPath
	default:
		return fmt.Errorf("unknown RPC session mode %q", s.mode)
	}
	return nil
}

func resolveRPCLocale(cmd *cobra.Command) (string, error) {
	if cmd.Flags().Changed("locale") {
		locale, ok := normalizeLocale(rpcLocale)
		if !ok {
			return "", fmt.Errorf("invalid --locale %q", rpcLocale)
		}
		return locale, nil
	}

	if raw, ok := os.LookupEnv("WITAN_LOCALE"); ok && strings.TrimSpace(raw) != "" {
		locale, valid := normalizeLocale(raw)
		if !valid {
			return "", fmt.Errorf("invalid WITAN_LOCALE %q", raw)
		}
		return locale, nil
	}

	if raw, ok := os.LookupEnv("LC_ALL"); ok && strings.TrimSpace(raw) != "" {
		locale, _ := normalizeLocale(raw)
		return locale, nil
	}

	for _, key := range []string{"LC_MESSAGES", "LANG"} {
		raw, ok := os.LookupEnv(key)
		if !ok || strings.TrimSpace(raw) == "" {
			continue
		}
		if locale, valid := normalizeLocale(raw); valid {
			return locale, nil
		}
	}

	return "", nil
}
