package cmd

import (
	"fmt"
	"net/http"
	"os"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/witanlabs/witan-cli/config"
)

const savedSessionSource = "saved session"

var authStatusJSON bool

var authStatusCmd = &cobra.Command{
	Use:   "status",
	Short: "Show active authentication status",
	Long: `Show which authentication credential Witan CLI would use right now.

Reports:
  - the active credential type and source
  - whether it validates successfully
  - the active organization when known
  - ignored lower-priority credentials

Examples:
  witan auth status
  witan auth status --json`,
	RunE: runAuthStatus,
}

type authStatusReport struct {
	Status             string                 `json:"status"`
	ActiveAuth         authCredentialReport   `json:"active_auth"`
	IgnoredCredentials []authCredentialReport `json:"ignored_credentials,omitempty"`
	Error              string                 `json:"error,omitempty"`
	Hint               string                 `json:"hint,omitempty"`
}

type authCredentialReport struct {
	Type                  string `json:"type"`
	Source                string `json:"source,omitempty"`
	CredentialFingerprint string `json:"credential_fingerprint,omitempty"`
	OrgID                 string `json:"org_id,omitempty"`
	UserEmail             string `json:"user_email,omitempty"`
	Validation            string `json:"validation,omitempty"`
	ValidationError       string `json:"validation_error,omitempty"`
}

func init() {
	authStatusCmd.SilenceUsage = true
	authStatusCmd.Flags().BoolVar(&authStatusJSON, "json", false, "Output raw JSON authentication status")
	authCmd.AddCommand(authStatusCmd)
}

func runAuthStatus(cmd *cobra.Command, args []string) error {
	report := inspectAuthStatus()
	if authStatusJSON {
		return jsonPrintTo(cmd.OutOrStdout(), report)
	}
	printAuthStatus(cmd, report)
	return nil
}

func inspectAuthStatus() authStatusReport {
	report := authStatusReport{
		ActiveAuth: authCredentialReport{Type: "none"},
	}

	flagAPIKey := apiKey
	envAPIKey := os.Getenv("WITAN_API_KEY")
	cfg, cfgErr := config.Load()

	switch {
	case flagAPIKey != "":
		report.ActiveAuth = inspectAPIKeyCredential(flagAPIKey, "--api-key", cfg, cfgErr == nil)
		if envAPIKey != "" {
			report.IgnoredCredentials = append(report.IgnoredCredentials, ignoredAPIKeyCredential(envAPIKey, "WITAN_API_KEY"))
		}
		if cfgErr == nil && cfg.SessionToken != "" {
			report.IgnoredCredentials = append(report.IgnoredCredentials, ignoredSessionCredential(cfg))
		}
		finalizeAuthStatus(&report)
		return report
	case envAPIKey != "":
		report.ActiveAuth = inspectAPIKeyCredential(envAPIKey, "WITAN_API_KEY", cfg, cfgErr == nil)
		if cfgErr == nil && cfg.SessionToken != "" {
			report.IgnoredCredentials = append(report.IgnoredCredentials, ignoredSessionCredential(cfg))
		}
		finalizeAuthStatus(&report)
		return report
	case cfgErr != nil:
		report.Status = "unauthenticated"
		report.Error = fmt.Sprintf("loading auth config: %v", cfgErr)
		report.Hint = "run `witan auth login` or set `WITAN_API_KEY`"
		return report
	case cfg.SessionToken != "":
		report.ActiveAuth = inspectSessionCredential(cfg.SessionToken, cfg.SessionOrgID)
		finalizeAuthStatus(&report)
		return report
	default:
		report.Status = "unauthenticated"
		report.Hint = "run `witan auth login` or set `WITAN_API_KEY`"
		return report
	}
}

func inspectAPIKeyCredential(rawKey, source string, cfg config.Config, hasConfig bool) authCredentialReport {
	report := authCredentialReport{
		Type:                  "api_key",
		Source:                source,
		CredentialFingerprint: config.HashAPIKey(rawKey),
	}

	cachedOrgID := ""
	if hasConfig {
		cachedOrgID = cfg.OrgIDForAPIKey(rawKey)
	}

	orgs, err := listOrgsByAPIKey(resolveManagementAPIURL(), rawKey)
	if err == nil {
		if len(orgs) == 0 {
			report.Validation = "invalid"
			report.ValidationError = "no organizations found for this API key"
			return report
		}
		report.Validation = "ok"
		if cachedOrgID != "" {
			report.OrgID = cachedOrgID
		} else {
			report.OrgID = orgs[0].ID
		}
		return report
	}

	report.Validation, report.ValidationError = classifyAuthValidationError(err)
	if report.Validation == "unknown" && cachedOrgID != "" {
		report.OrgID = cachedOrgID
	}
	return report
}

func inspectSessionCredential(sessionToken, sessionOrgID string) authCredentialReport {
	report := authCredentialReport{
		Type:   "session",
		Source: savedSessionSource,
		OrgID:  sessionOrgID,
	}

	if _, err := exchangeSessionForJWT(resolveManagementAPIURL(), sessionToken); err != nil {
		report.Validation, report.ValidationError = classifyAuthValidationError(err)
		return report
	}

	report.Validation = "ok"

	httpClient := &http.Client{Timeout: 10 * time.Second}
	session, err := getSession(httpClient, resolveManagementAPIURL(), sessionToken)
	if err == nil {
		report.UserEmail = strings.TrimSpace(session.User.Email)
	}

	return report
}

func ignoredAPIKeyCredential(rawKey, source string) authCredentialReport {
	return authCredentialReport{
		Type:                  "api_key",
		Source:                source,
		CredentialFingerprint: config.HashAPIKey(rawKey),
	}
}

func ignoredSessionCredential(cfg config.Config) authCredentialReport {
	return authCredentialReport{
		Type:   "session",
		Source: savedSessionSource,
		OrgID:  cfg.SessionOrgID,
	}
}

func finalizeAuthStatus(report *authStatusReport) {
	switch report.ActiveAuth.Validation {
	case "ok":
		report.Status = "authenticated"
	case "invalid":
		report.Status = "unauthenticated"
	case "unknown":
		report.Status = "unknown"
	default:
		report.Status = "unauthenticated"
	}

	if report.Hint != "" {
		return
	}

	switch report.ActiveAuth.Type {
	case "api_key":
		if report.ActiveAuth.Validation == "invalid" {
			if hasIgnoredCredential(report.IgnoredCredentials, "session") {
				report.Hint = "remove or replace the active API key to use the saved session"
			} else {
				report.Hint = "update the API key or run `witan auth login`"
			}
		}
	case "session":
		if report.ActiveAuth.Validation == "invalid" {
			report.Hint = "run `witan auth login` again"
		}
	}
}

func hasIgnoredCredential(credentials []authCredentialReport, kind string) bool {
	for _, credential := range credentials {
		if credential.Type == kind {
			return true
		}
	}
	return false
}

func classifyAuthValidationError(err error) (string, string) {
	if err == nil {
		return "ok", ""
	}

	msg := err.Error()
	switch {
	case strings.Contains(msg, "HTTP 401"), strings.Contains(msg, "HTTP 403"), strings.Contains(msg, "no organizations found for this API key"):
		return "invalid", msg
	default:
		return "unknown", msg
	}
}

func printAuthStatus(cmd *cobra.Command, report authStatusReport) {
	out := cmd.OutOrStdout()

	fmt.Fprintf(out, "Status: %s\n", report.Status)
	fmt.Fprintf(out, "Active auth: %s\n", humanAuthType(report.ActiveAuth.Type))

	if report.ActiveAuth.Source != "" {
		fmt.Fprintf(out, "Source: %s\n", report.ActiveAuth.Source)
	}
	if report.ActiveAuth.CredentialFingerprint != "" {
		fmt.Fprintf(out, "Credential: sha256:%s\n", shortFingerprint(report.ActiveAuth.CredentialFingerprint))
	}
	if report.ActiveAuth.UserEmail != "" {
		fmt.Fprintf(out, "User: %s\n", report.ActiveAuth.UserEmail)
	}
	if report.ActiveAuth.OrgID != "" {
		fmt.Fprintf(out, "Org: %s\n", report.ActiveAuth.OrgID)
	}
	if report.ActiveAuth.Validation != "" {
		fmt.Fprintf(out, "Validation: %s", report.ActiveAuth.Validation)
		if report.ActiveAuth.ValidationError != "" {
			fmt.Fprintf(out, " (%s)", report.ActiveAuth.ValidationError)
		}
		fmt.Fprintln(out)
	}
	if report.Error != "" {
		fmt.Fprintf(out, "Error: %s\n", report.Error)
	}
	if len(report.IgnoredCredentials) > 0 {
		fmt.Fprintln(out, "Ignored credentials:")
		for _, credential := range report.IgnoredCredentials {
			fmt.Fprintf(out, "- %s", humanAuthType(credential.Type))
			if credential.Source != "" {
				fmt.Fprintf(out, " (%s", credential.Source)
				if credential.CredentialFingerprint != "" {
					fmt.Fprintf(out, ", sha256:%s", shortFingerprint(credential.CredentialFingerprint))
				}
				fmt.Fprint(out, ")")
			} else if credential.CredentialFingerprint != "" {
				fmt.Fprintf(out, " (sha256:%s)", shortFingerprint(credential.CredentialFingerprint))
			}
			if credential.OrgID != "" {
				fmt.Fprintf(out, " org=%s", credential.OrgID)
			}
			fmt.Fprintln(out)
		}
	}
	if report.Hint != "" {
		fmt.Fprintf(out, "Hint: %s\n", report.Hint)
	}
}

func humanAuthType(kind string) string {
	switch kind {
	case "api_key":
		return "api key"
	case "session":
		return "session"
	default:
		return kind
	}
}

func shortFingerprint(fingerprint string) string {
	if len(fingerprint) > 12 {
		return fingerprint[:12]
	}
	return fingerprint
}
