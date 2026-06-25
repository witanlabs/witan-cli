package cmd

import (
	"errors"
	"os"
	"testing"

	"github.com/witanlabs/witan-cli/client"
)

// silenceStderr redirects os.Stderr for the duration of fn so error-path tests
// don't spam test output.
func silenceStderr(t *testing.T, fn func()) {
	t.Helper()
	orig := os.Stderr
	devnull, _ := os.Open(os.DevNull)
	os.Stderr = devnull
	defer func() {
		os.Stderr = orig
		devnull.Close()
	}()
	fn()
}

func TestHandleSheetsOpError_needsFileAuthorization(t *testing.T) {
	var got error
	silenceStderr(t, func() {
		got = handleSheetsOpError(&client.APIError{StatusCode: 403, Code: "needs_file_authorization"}, "abc123", false)
	})
	var exitErr *ExitError
	if !errors.As(got, &exitErr) || exitErr.Code != authRequiredExitCode {
		t.Fatalf("got %v, want ExitError code %d", got, authRequiredExitCode)
	}
}

func TestHandleSheetsOpError_googleAuthRequired(t *testing.T) {
	var got error
	silenceStderr(t, func() {
		got = handleSheetsOpError(&client.APIError{StatusCode: 401, Code: "google_auth_required"}, "abc123", false)
	})
	var exitErr *ExitError
	if !errors.As(got, &exitErr) || exitErr.Code != authRequiredExitCode {
		t.Fatalf("got %v, want ExitError code %d", got, authRequiredExitCode)
	}
}

func TestHandleSheetsOpError_bare401SessionExpiredPassesThrough(t *testing.T) {
	// A 401 whose code is NOT google_auth_required means the Witan session
	// expired ('witan auth login'), not Google reconnect — must pass through.
	in := &client.APIError{StatusCode: 401, Code: "unauthorized"}
	got := handleSheetsOpError(in, "abc123", false)
	var exitErr *ExitError
	if errors.As(got, &exitErr) {
		t.Fatalf("bare 401 should not become an auth-required ExitError, got %v", got)
	}
	if got != in {
		t.Fatalf("expected passthrough of the original error, got %v", got)
	}
}

func TestHandleSheetsOpError_passthrough(t *testing.T) {
	in := &client.APIError{StatusCode: 500, Code: "internal"}
	if got := handleSheetsOpError(in, "abc123", false); got != in {
		t.Fatalf("expected passthrough of non-auth error, got %v", got)
	}

	plain := errors.New("network down")
	if got := handleSheetsOpError(plain, "abc123", false); got != plain {
		t.Fatalf("expected passthrough of non-APIError, got %v", got)
	}

	if got := handleSheetsOpError(nil, "abc123", false); got != nil {
		t.Fatalf("expected nil for nil error, got %v", got)
	}
}

func TestSheetsRPCInitFailure_authExitCode(t *testing.T) {
	var got error
	silenceStderr(t, func() {
		got = sheetsRPCInitFailure(&sheetsRPCInitResponse{Code: "needs_file_authorization", Message: "x"}, "abc123")
	})
	var exitErr *ExitError
	if !errors.As(got, &exitErr) || exitErr.Code != authRequiredExitCode {
		t.Fatalf("got %v, want ExitError code %d", got, authRequiredExitCode)
	}
}

func TestSheetsRPCInitFailure_otherPassesToFormatter(t *testing.T) {
	got := sheetsRPCInitFailure(&sheetsRPCInitResponse{Code: "INVALID_INIT", Message: "bad fields"}, "abc123")
	var exitErr *ExitError
	if errors.As(got, &exitErr) {
		t.Fatalf("non-auth code should not be an ExitError, got %v", got)
	}
	if got == nil || got.Error() == "" {
		t.Fatalf("expected a descriptive error, got %v", got)
	}
}
