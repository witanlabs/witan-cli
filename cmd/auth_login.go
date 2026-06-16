package cmd

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"os/signal"
	"runtime"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/config"
	"golang.org/x/term"
)

var loginCmd = &cobra.Command{
	Use:   "login",
	Short: "Authenticate with Witan via browser",
	Long: `Start browser-based sign-in with a one-time device code.

What happens:
  1. The CLI prints a one-time code and opens the verification page.
  2. After approval, the CLI exchanges the device code for a session token.
  3. If needed, you are prompted to choose an active organization.
  4. The session is saved locally for future commands.

When stdin is not a terminal (e.g. driven by an agent), or with --json, login
runs non-interactively: it prints the verification URL and one-time code up
front, does not open a browser, and polls to completion in the same process.
Hand the URL/code to a human on another device.

With multiple organizations, select one non-interactively via --org <id> or
WITAN_ORG. If neither is set in non-interactive mode, the organization list is
emitted and the command exits with code 3 (the session is saved, so a re-run
with --org finishes without re-authenticating).

In --json mode, stdout is newline-delimited JSON (one object per line), each
tagged with a "type": device_authorization (the verification URL/code),
org_selection_required (exit code 3), or login_complete.

For non-session, fully unattended use, prefer --api-key or WITAN_API_KEY.

Example:
  witan auth login
  witan auth login --json --org org_123`,
	RunE: runLogin,
}

var (
	loginJSON      bool
	loginNoBrowser bool
	loginOrg       string
)

func init() {
	loginCmd.SilenceUsage = true
	loginCmd.Flags().BoolVar(&loginJSON, "json", false, "Emit machine-readable JSONL events (device_authorization, org_selection_required, login_complete) and run non-interactively")
	loginCmd.Flags().BoolVar(&loginNoBrowser, "no-browser", false, "Do not attempt to open a browser")
	loginCmd.Flags().StringVar(&loginOrg, "org", "", "Organization ID to select (env: WITAN_ORG)")
	authCmd.AddCommand(loginCmd)
}

// stdinIsTTY reports whether stdin is an interactive terminal. When it is not
// (a pipe or /dev/null, as when an agent or daemon spawns the process), login
// must avoid blocking reads from stdin. A character-device check is not enough
// here: /dev/null is itself a character device, so this uses a real isatty(fd)
// check on stdin.
func stdinIsTTY() bool {
	return term.IsTerminal(int(os.Stdin.Fd()))
}

// resolveLoginOrg returns the org selection from --org or WITAN_ORG.
func resolveLoginOrg() string {
	if loginOrg != "" {
		return loginOrg
	}
	return os.Getenv("WITAN_ORG")
}

// canResumeOrgSelection reports whether a saved session is an incomplete
// multi-org login — a token saved without an org, as left behind when a
// non-interactive run exited needing --org — that a re-run with an org
// preference can finish without re-authenticating. A completed session (org
// already set) is deliberately excluded: `auth login` must always
// re-authenticate, so it never silently reuses an existing active session.
func canResumeOrgSelection(cfg config.Config, nonInteractive bool, orgPref string) bool {
	return nonInteractive && orgPref != "" && cfg.SessionToken != "" && cfg.SessionOrgID == ""
}

func orgContains(orgs []orgEntry, id string) bool {
	for _, o := range orgs {
		if o.ID == id {
			return true
		}
	}
	return false
}

type deviceCodeResponse struct {
	DeviceCode              string `json:"device_code"`
	UserCode                string `json:"user_code"`
	VerificationURI         string `json:"verification_uri"`
	VerificationURIComplete string `json:"verification_uri_complete"`
	ExpiresIn               int    `json:"expires_in"`
	Interval                int    `json:"interval"`
}

type tokenResponse struct {
	AccessToken string `json:"access_token"`
	TokenType   string `json:"token_type"`
}

type tokenErrorResponse struct {
	Error            string `json:"error"`
	ErrorDescription string `json:"error_description"`
}

type sessionResponse struct {
	User struct {
		Email string `json:"email"`
	} `json:"user"`
}

func runLogin(cmd *cobra.Command, args []string) error {
	mgmtURL := resolveManagementAPIURL()
	httpClient := &http.Client{Timeout: 30 * time.Second}

	nonInteractive := loginJSON || !stdinIsTTY()
	orgPref := resolveLoginOrg()

	// Non-interactive re-run to finish org selection: if a prior run saved a
	// session token but no org (an incomplete multi-org login that exited needing
	// --org) and an org is now specified, complete login by reusing that token —
	// this avoids forcing the human to approve a second time. If the saved token
	// is no longer valid, fall through to a fresh device-code flow.
	if cfg, err := config.Load(); err == nil && canResumeOrgSelection(cfg, nonInteractive, orgPref) {
		err := completeLogin(httpClient, mgmtURL, cfg.SessionToken, orgPref, nonInteractive)
		if err == nil {
			return nil
		}
		if !isInvalidSavedSessionError(err) {
			return err
		}
		// invalid/expired saved session — proceed with a fresh login below
	}

	// Step 1: Request device code
	body, _ := json.Marshal(map[string]string{"client_id": "witan-cli"})
	req, err := http.NewRequest("POST", mgmtURL+"/v0/auth/device/code", bytes.NewReader(body))
	if err != nil {
		return fmt.Errorf("failed to request device code: %w", err)
	}
	req.Header.Set("Content-Type", "application/json")
	setCLIUserAgent(req)
	resp, err := httpClient.Do(req)
	if err != nil {
		return fmt.Errorf("failed to request device code: %w", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		respBody, _ := io.ReadAll(resp.Body)
		return fmt.Errorf("failed to request device code (HTTP %d): %s", resp.StatusCode, string(respBody))
	}

	var dcResp deviceCodeResponse
	if err := json.NewDecoder(resp.Body).Decode(&dcResp); err != nil {
		return fmt.Errorf("failed to parse device code response: %w", err)
	}

	// Step 2: Display code and open browser
	displayCode := dcResp.UserCode
	if len(displayCode) >= 8 {
		displayCode = displayCode[:4] + "-" + displayCode[4:]
	}
	if nonInteractive {
		emitHandoff(&dcResp, displayCode)
	} else {
		fmt.Fprintf(os.Stderr, "! First, copy your one-time code: %s\n", displayCode)
		fmt.Fprintf(os.Stderr, "Press Enter to open %s in your browser...", dcResp.VerificationURI)

		// Wait for Enter
		buf := make([]byte, 1)
		os.Stdin.Read(buf)

		if !loginNoBrowser {
			if err := openBrowser(dcResp.VerificationURI); err != nil {
				fmt.Fprintf(os.Stderr, "Could not open browser. Please visit:\n  %s\n", dcResp.VerificationURI)
			}
		}
	}

	// Step 3: Poll for token
	ctx, cancel := signal.NotifyContext(context.Background(), os.Interrupt)
	defer cancel()

	interval := time.Duration(dcResp.Interval) * time.Second
	if interval < 5*time.Second {
		interval = 5 * time.Second
	}

	fmt.Fprintf(os.Stderr, "Waiting for authorization...\n")

	deadline := time.Now().Add(time.Duration(dcResp.ExpiresIn) * time.Second)

	var sessionToken string
	for {
		select {
		case <-ctx.Done():
			return fmt.Errorf("interrupted")
		case <-time.After(interval):
		}

		if time.Now().After(deadline) {
			return fmt.Errorf("code expired, please run 'witan auth login' again")
		}

		token, done, err := pollToken(httpClient, mgmtURL, dcResp.DeviceCode, &interval)
		if err != nil {
			return err
		}
		if done {
			sessionToken = token
			break
		}
	}

	// Steps 4 & 5: resolve session, select org, and save config.
	return completeLogin(httpClient, mgmtURL, sessionToken, orgPref, nonInteractive)
}

// emitHandoff prints the device-code verification payload for a human on
// another device. In --json mode it writes a machine-readable object to stdout;
// otherwise it prints a human-readable prompt to stderr. It never reads stdin
// or opens a browser.
func emitHandoff(dc *deviceCodeResponse, displayCode string) {
	if loginJSON {
		jsonlPrint(map[string]any{
			"type":                      "device_authorization",
			"verification_uri":          dc.VerificationURI,
			"verification_uri_complete": dc.VerificationURIComplete,
			"user_code":                 dc.UserCode,
			"expires_in":                dc.ExpiresIn,
		})
		return
	}
	target := dc.VerificationURIComplete
	if target == "" {
		target = dc.VerificationURI
	}
	fmt.Fprintf(os.Stderr, "To sign in, open this URL in a browser:\n  %s\n", target)
	fmt.Fprintf(os.Stderr, "and enter the code: %s\n", displayCode)
}

// completeLogin exchanges a freshly minted session token for the user's orgs,
// selects one, and saves the config. The sessionToken is assumed valid; an
// HTTP 401/403 surfaces as an invalid-session error so callers reusing a saved
// token can fall back to a fresh login.
func completeLogin(client *http.Client, mgmtURL, sessionToken, orgPref string, nonInteractive bool) error {
	session, err := getSession(client, mgmtURL, sessionToken)
	if err != nil {
		return fmt.Errorf("failed to get session: %w", err)
	}
	email := session.User.Email

	jwt, err := exchangeSessionForJWT(mgmtURL, sessionToken)
	if err != nil {
		return fmt.Errorf("failed to exchange session for JWT: %w", err)
	}

	orgs, err := listOrgsByJWT(mgmtURL, jwt)
	if err != nil {
		return fmt.Errorf("failed to list organizations: %w", err)
	}

	selectedOrgID, err := selectOrg(orgs, orgPref, sessionToken, nonInteractive)
	if err != nil {
		return err
	}

	// Save config
	if err := config.Save(config.Config{
		SessionToken: sessionToken,
		SessionOrgID: selectedOrgID,
	}); err != nil {
		return fmt.Errorf("failed to save config: %w", err)
	}

	emitLoginComplete(email, selectedOrgID)
	if email != "" {
		fmt.Fprintf(os.Stderr, "\u2713 Logged in as %s\n", email)
	} else {
		fmt.Fprintf(os.Stderr, "\u2713 Logged in\n")
	}

	return nil
}

// emitLoginComplete writes the terminal success event in --json mode so a
// machine consumer reading stdout has a structured signal (and the resulting
// org) rather than only an exit code. It is a no-op outside --json; the
// human-readable confirmation is always printed to stderr by the caller.
func emitLoginComplete(email, orgID string) {
	if !loginJSON {
		return
	}
	jsonlPrint(map[string]any{
		"type":   "login_complete",
		"email":  email,
		"org_id": orgID,
	})
}

// selectOrg chooses the active organization. A non-empty orgPref must match one
// of the user's orgs. With multiple orgs and no preference: in non-interactive
// mode the org list is emitted, the session token is saved (so a re-run with
// --org can finish without re-authenticating), and an &ExitError{Code: 3} is
// returned; interactively, the user is prompted.
func selectOrg(orgs []orgEntry, orgPref, sessionToken string, nonInteractive bool) (string, error) {
	if orgPref != "" {
		if !orgContains(orgs, orgPref) {
			return "", fmt.Errorf("organization %q not found among your organizations", orgPref)
		}
		return orgPref, nil
	}

	switch len(orgs) {
	case 0:
		return "", fmt.Errorf("no organizations found \u2014 contact your administrator")
	case 1:
		return orgs[0].ID, nil
	default:
		if nonInteractive {
			// Save the session so a re-run with --org finishes without re-auth.
			if err := config.Save(config.Config{SessionToken: sessionToken}); err != nil {
				return "", fmt.Errorf("failed to save config: %w", err)
			}
			emitOrgChoices(orgs)
			return "", &ExitError{Code: 3}
		}
		return promptOrg(orgs)
	}
}

// emitOrgChoices reports the available organizations for non-interactive
// selection: JSON to stdout under --json, otherwise a list to stderr.
func emitOrgChoices(orgs []orgEntry) {
	if loginJSON {
		jsonlPrint(map[string]any{
			"type":          "org_selection_required",
			"organizations": orgs,
		})
		return
	}
	fmt.Fprintf(os.Stderr, "Multiple organizations available. Re-run with --org <id> (or set WITAN_ORG):\n")
	for _, org := range orgs {
		fmt.Fprintf(os.Stderr, "  %s  %s\n", org.ID, org.Name)
	}
}

// promptOrg interactively asks the user to choose an organization from a list.
func promptOrg(orgs []orgEntry) (string, error) {
	fmt.Fprintf(os.Stderr, "? Select an organization:\n")
	for i, org := range orgs {
		fmt.Fprintf(os.Stderr, "  [%d] %s\n", i+1, org.Name)
	}
	for {
		fmt.Fprintf(os.Stderr, "Choice: ")
		var choice int
		_, err := fmt.Fscan(os.Stdin, &choice)
		if errors.Is(err, io.EOF) {
			// stdin closed (e.g. non-terminal that slipped past detection):
			// fail instead of looping forever on repeated EOF.
			return "", fmt.Errorf("no organization selected (stdin closed); re-run with --org <id> or WITAN_ORG")
		}
		if err != nil || choice < 1 || choice > len(orgs) {
			fmt.Fprintf(os.Stderr, "Invalid choice, enter a number between 1 and %d.\n", len(orgs))
			// drain the rest of the line on bad input
			var discard string
			fmt.Fscanln(os.Stdin, &discard)
			continue
		}
		return orgs[choice-1].ID, nil
	}
}

func pollToken(client *http.Client, mgmtURL, deviceCode string, interval *time.Duration) (string, bool, error) {
	body, _ := json.Marshal(map[string]string{
		"grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
		"device_code": deviceCode,
		"client_id":   "witan-cli",
	})

	req, err := http.NewRequest("POST", mgmtURL+"/v0/auth/device/token", bytes.NewReader(body))
	if err != nil {
		return "", false, fmt.Errorf("failed to poll for token: %w", err)
	}
	req.Header.Set("Content-Type", "application/json")
	setCLIUserAgent(req)

	resp, err := client.Do(req)
	if err != nil {
		return "", false, fmt.Errorf("failed to poll for token: %w", err)
	}
	defer resp.Body.Close()

	respBody, _ := io.ReadAll(resp.Body)

	if resp.StatusCode == http.StatusOK {
		var tr tokenResponse
		if err := json.Unmarshal(respBody, &tr); err != nil {
			return "", false, fmt.Errorf("failed to parse token response: %w", err)
		}
		return tr.AccessToken, true, nil
	}

	var errResp tokenErrorResponse
	if err := json.Unmarshal(respBody, &errResp); err != nil {
		return "", false, fmt.Errorf("unexpected response (HTTP %d): %s", resp.StatusCode, string(respBody))
	}

	switch errResp.Error {
	case "authorization_pending":
		return "", false, nil
	case "slow_down":
		*interval += 5 * time.Second
		return "", false, nil
	case "expired_token":
		return "", false, fmt.Errorf("code expired, please run 'witan auth login' again")
	case "access_denied":
		return "", false, fmt.Errorf("login denied by user")
	default:
		return "", false, fmt.Errorf("authorization failed: %s — %s", errResp.Error, errResp.ErrorDescription)
	}
}

func getSession(client *http.Client, mgmtURL, token string) (*sessionResponse, error) {
	req, err := http.NewRequest("GET", mgmtURL+"/v0/auth/get-session", nil)
	if err != nil {
		return nil, fmt.Errorf("invalid management API URL: %w", err)
	}
	setCLIUserAgent(req)
	req.Header.Set("Authorization", "Bearer "+token)
	resp, err := client.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("HTTP %d", resp.StatusCode)
	}
	var s sessionResponse
	if err := json.NewDecoder(resp.Body).Decode(&s); err != nil {
		return nil, err
	}
	return &s, nil
}

func openBrowser(url string) error {
	var cmd *exec.Cmd
	switch runtime.GOOS {
	case "darwin":
		cmd = exec.Command("open", url)
	case "linux":
		cmd = exec.Command("xdg-open", url)
	case "windows":
		cmd = exec.Command("cmd", "/c", "start", url)
	default:
		return fmt.Errorf("unsupported platform %s", runtime.GOOS)
	}
	return cmd.Start()
}
