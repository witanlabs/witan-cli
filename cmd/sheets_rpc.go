package cmd

import (
	"bufio"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/coder/websocket"
	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/client"
)

var (
	sheetsRPCLocale string
	sheetsRPCCreate bool
	sheetsRPCTitle  string
)

var sheetsRPCCmd = &cobra.Command{
	Use:   "rpc [<spreadsheet>]",
	Short: "Run newline-delimited RPC over stdio for Google Sheets",
	Long: `Run newline-delimited RPC over stdio for Google Sheets.

The command opens a WebSocket connection to the Google Sheets RPC endpoint, then
relays one JSON object per input line. Stdout contains one JSON response per
request. Stderr is reserved for CLI diagnostics.

Spreadsheet reference:
  - Full URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  - Short form: gs://SPREADSHEET_ID
  - Create mode: omit the argument, pass new or gs://new, or use --create

The spreadsheet is bound during the WebSocket init handshake (not in the URL path).
Use --create to start a new spreadsheet session; the init response includes the
new spreadsheet_id and url on stderr.

Input shape:
  {"id":"1","op":"listSheets","args":{}}
  {"id":"2","op":"readRange","args":{"address":"Sheet1!A1:B10"}}
  {"id":"3","op":"setCells","args":{"cells":[...]}}

Blocked operations: open, create, close, health, evaluatescript (use 'gsheets exec').

Note: Changes auto-persist to Google Sheets. The 'save' operation returns success
immediately.

Examples:
  witan gsheets rpc gs://SPREADSHEET_ID
  echo '{"id":"1","op":"listSheets","args":{}}' | witan gsheets rpc gs://ID
  witan gsheets rpc --create --title "Q4 Model"
  echo '{"id":"1","op":"setCells","args":{"cells":[{"address":"Sheet1!A1","value":"Hi"}]}}' | witan gsheets rpc --create`,
	Args: validateSheetsRPCArgs,
	RunE: runSheetsRPC,
}

func init() {
	sheetsRPCCmd.SilenceUsage = true
	sheetsRPCCmd.Flags().StringVar(&sheetsRPCLocale, "locale", "", "Execution locale (env: WITAN_LOCALE; otherwise LC_ALL / LC_MESSAGES / LANG)")
	sheetsRPCCmd.Flags().BoolVar(&sheetsRPCCreate, "create", false, "Create a new Google Sheet instead of opening an existing one")
	sheetsRPCCmd.Flags().StringVar(&sheetsRPCTitle, "title", "", "Title for a newly created spreadsheet (create mode only, max 1000 characters)")
	gsheetsCmd.AddCommand(sheetsRPCCmd)
}

func validateSheetsRPCArgs(_ *cobra.Command, args []string) error {
	return validateSheetsOpenOrCreateArgs(args, sheetsRPCCreate)
}

func runSheetsRPC(cmd *cobra.Command, args []string) error {
	create := resolveSheetsCreate(sheetsRPCCreate, args)

	if err := validateSheetsTitle(sheetsRPCTitle, create); err != nil {
		return err
	}
	if sheetsRPCTitle != "" && !create {
		return fmt.Errorf("--title can only be used with --create or spreadsheet reference new")
	}

	locale, err := resolveLocale(cmd, "locale", sheetsRPCLocale, true, false)
	if err != nil {
		return err
	}

	auth, err := requireSheetsAuth()
	if err != nil {
		return err
	}

	var spreadsheetID string
	if !create {
		spreadsheetID = client.ExtractSpreadsheetID(args[0])
	}

	session, err := openSheetsRPCSession(cmd.Context(), auth.Client, sheetsRPCConnectParams{
		Create:        create,
		SpreadsheetID: spreadsheetID,
		Title:         sheetsRPCTitle,
		Locale:        locale,
	})
	if err != nil {
		return err
	}
	defer session.close()

	if create && !gsheetsJSONOutput {
		title := session.title
		if title == "" {
			title = sheetsRPCTitle
		}
		outputSheetsCreateHints(session.spreadsheetID, session.url, title)
	}

	return relaySheetsRPCStdio(cmd.Context(), session, os.Stdin, os.Stdout)
}

type sheetsRPCConnectParams struct {
	Create        bool
	SpreadsheetID string
	Title         string
	Locale        string
}

type sheetsRPCSession struct {
	conn          *websocket.Conn
	spreadsheetID string
	url           string
	title         string
}

type sheetsRPCInitMessage struct {
	Type          string `json:"type"`
	ID            string `json:"id"`
	Locale        string `json:"locale,omitempty"`
	SpreadsheetID string `json:"spreadsheet_id,omitempty"`
	Create        bool   `json:"create,omitempty"`
	Title         string `json:"title,omitempty"`
}

type sheetsRPCInitResponse struct {
	ID            string `json:"id"`
	Ok            bool   `json:"ok"`
	Type          string `json:"type"`
	SpreadsheetID string `json:"spreadsheet_id"`
	URL           string `json:"url"`
	Title         string `json:"title"`
	Code          string `json:"code"`
	Message       string `json:"message"`
}

func openSheetsRPCSession(ctx context.Context, c *client.Client, params sheetsRPCConnectParams) (*sheetsRPCSession, error) {
	wsURL, err := c.GSheetsRPCWebSocketURL()
	if err != nil {
		return nil, err
	}

	conn, err := dialRPCWebSocket(ctx, wsURL, c.APIKey, cliUserAgent())
	if err != nil {
		return nil, err
	}
	conn.SetReadLimit(rpcReadLimit)

	initMsg := sheetsRPCInitMessage{
		Type:   "init",
		ID:     "witan-init-1",
		Locale: params.Locale,
	}
	if params.Create {
		initMsg.Create = true
		if params.Title != "" {
			initMsg.Title = params.Title
		}
	} else {
		initMsg.SpreadsheetID = params.SpreadsheetID
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

	var initResp sheetsRPCInitResponse
	if err := json.Unmarshal(resp, &initResp); err != nil {
		conn.CloseNow()
		return nil, fmt.Errorf("parsing init response: %w", err)
	}
	if err := formatSheetsRPCInitError(&initResp); err != nil {
		conn.CloseNow()
		return nil, err
	}

	return &sheetsRPCSession{
		conn:          conn,
		spreadsheetID: initResp.SpreadsheetID,
		url:           initResp.URL,
		title:         initResp.Title,
	}, nil
}

func formatSheetsRPCInitError(resp *sheetsRPCInitResponse) error {
	if resp.Ok {
		if resp.SpreadsheetID == "" {
			return fmt.Errorf("init response missing spreadsheet_id")
		}
		return nil
	}
	if resp.Code != "" {
		switch resp.Code {
		case "google_auth_required":
			return fmt.Errorf("Google Sheets requires authorization. Run 'witan gsheets connect' to enable access.")
		case "google_sheets_not_found":
			return fmt.Errorf("spreadsheet not found or not shared with your account")
		case "google_sheets_forbidden":
			return fmt.Errorf("you don't have permission to access this spreadsheet")
		case "INVALID_INIT":
			return fmt.Errorf("%s: %s", resp.Code, resp.Message)
		}
		return fmt.Errorf("%s: %s", resp.Code, resp.Message)
	}
	if resp.Message != "" {
		return fmt.Errorf("initializing RPC session: %s", resp.Message)
	}
	return fmt.Errorf("initializing RPC session failed")
}

func (s *sheetsRPCSession) close() {
	if s.conn != nil {
		_ = s.conn.Close(websocket.StatusNormalClosure, "")
	}
}

func relaySheetsRPCStdio(ctx context.Context, session *sheetsRPCSession, stdin io.Reader, stdout io.Writer) error {
	scanner := bufio.NewScanner(stdin)
	scanner.Buffer(make([]byte, 0, 64*1024), rpcReadLimit)

	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "" {
			continue
		}

		if err := session.conn.Write(ctx, websocket.MessageText, []byte(line)); err != nil {
			return fmt.Errorf("sending RPC message: %w", err)
		}

		msgType, rawResp, err := session.conn.Read(ctx)
		if err != nil {
			return fmt.Errorf("reading RPC response: %w", err)
		}
		if msgType != websocket.MessageText {
			return fmt.Errorf("reading RPC response: expected text frame, got %v", msgType)
		}

		if _, err := stdout.Write(rawResp); err != nil {
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
