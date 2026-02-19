package cmd

import (
	"bytes"
	"context"
	"encoding/json"
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

For non-interactive environments, use --api-key or WITAN_API_KEY.

Example:
  witan auth login`,
	RunE: runLogin,
}

func init() {
	loginCmd.SilenceUsage = true
	authCmd.AddCommand(loginCmd)
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
	Session struct {
		ActiveOrganizationId string `json:"activeOrganizationId"`
	} `json:"session"`
	User struct {
		Email string `json:"email"`
	} `json:"user"`
}

type listOrgsResponse []struct {
	ID   string `json:"id"`
	Name string `json:"name"`
}

func runLogin(cmd *cobra.Command, args []string) error {
	mgmtURL := resolveManagementAPIURL()
	httpClient := &http.Client{Timeout: 30 * time.Second}

	// Step 1: Request device code
	body, _ := json.Marshal(map[string]string{"client_id": "witan-cli"})
	resp, err := httpClient.Post(mgmtURL+"/v0/auth/device/code", "application/json", bytes.NewReader(body))
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
	fmt.Fprintf(os.Stderr, "! First, copy your one-time code: %s\n", displayCode)
	fmt.Fprintf(os.Stderr, "Press Enter to open %s in your browser...", dcResp.VerificationURI)

	// Wait for Enter
	buf := make([]byte, 1)
	os.Stdin.Read(buf)

	if err := openBrowser(dcResp.VerificationURI); err != nil {
		fmt.Fprintf(os.Stderr, "Could not open browser. Please visit:\n  %s\n", dcResp.VerificationURI)
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

	// Step 4: Fetch session (and capture email for later display)
	session, err := getSession(httpClient, mgmtURL, sessionToken)
	if err != nil {
		return fmt.Errorf("failed to get session: %w", err)
	}
	email := session.User.Email

	if session.Session.ActiveOrganizationId == "" {
		orgs, err := listOrganizations(httpClient, mgmtURL, sessionToken)
		if err != nil {
			return fmt.Errorf("failed to list organizations: %w", err)
		}

		var selectedOrgID string
		switch len(orgs) {
		case 0:
			return fmt.Errorf("no organizations found — contact your administrator")
		case 1:
			selectedOrgID = orgs[0].ID
		default:
			fmt.Fprintf(os.Stderr, "? Select an organization:\n")
			for i, org := range orgs {
				fmt.Fprintf(os.Stderr, "  [%d] %s\n", i+1, org.Name)
			}
			for {
				fmt.Fprintf(os.Stderr, "Choice: ")
				var choice int
				if _, err := fmt.Fscan(os.Stdin, &choice); err != nil || choice < 1 || choice > len(orgs) {
					fmt.Fprintf(os.Stderr, "Invalid choice, enter a number between 1 and %d.\n", len(orgs))
					// drain the rest of the line on bad input
					var discard string
					fmt.Fscanln(os.Stdin, &discard)
					continue
				}
				selectedOrgID = orgs[choice-1].ID
				break
			}
		}

		if err := setActiveOrganization(httpClient, mgmtURL, sessionToken, selectedOrgID); err != nil {
			return fmt.Errorf("failed to set active organization: %w", err)
		}
	}

	// Step 5: Save config
	if err := config.Save(config.Config{SessionToken: sessionToken}); err != nil {
		return fmt.Errorf("failed to save config: %w", err)
	}

	if email != "" {
		fmt.Fprintf(os.Stderr, "\u2713 Logged in as %s\n", email)
	} else {
		fmt.Fprintf(os.Stderr, "\u2713 Logged in\n")
	}

	return nil
}

func pollToken(client *http.Client, mgmtURL, deviceCode string, interval *time.Duration) (string, bool, error) {
	body, _ := json.Marshal(map[string]string{
		"grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
		"device_code": deviceCode,
		"client_id":   "witan-cli",
	})

	resp, err := client.Post(mgmtURL+"/v0/auth/device/token", "application/json", bytes.NewReader(body))
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

func listOrganizations(client *http.Client, mgmtURL, token string) (listOrgsResponse, error) {
	req, err := http.NewRequest("GET", mgmtURL+"/v0/auth/organization/list", nil)
	if err != nil {
		return nil, fmt.Errorf("invalid management API URL: %w", err)
	}
	req.Header.Set("Authorization", "Bearer "+token)
	resp, err := client.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("HTTP %d", resp.StatusCode)
	}
	var orgs listOrgsResponse
	if err := json.NewDecoder(resp.Body).Decode(&orgs); err != nil {
		return nil, err
	}
	return orgs, nil
}

func setActiveOrganization(client *http.Client, mgmtURL, token, orgID string) error {
	body, _ := json.Marshal(map[string]string{"organizationId": orgID})
	req, err := http.NewRequest("POST", mgmtURL+"/v0/auth/organization/set-active", bytes.NewReader(body))
	if err != nil {
		return fmt.Errorf("invalid management API URL: %w", err)
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")
	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("HTTP %d", resp.StatusCode)
	}
	return nil
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
