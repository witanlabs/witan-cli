package cmd

import (
	"context"
	"fmt"
	"os"
	"time"
)

// authRequiredExitCode signals that a Google authorization step is needed
// (account connect or per-sheet authorize) before an operation can proceed.
// Agents branch on this exit code instead of parsing messages.
const authRequiredExitCode = 3

// isInteractive reports whether the CLI is attached to a terminal on both
// stdin and stdout. Browser-based flows (connect, authorize) only open a
// browser and block-poll when interactive.
func isInteractive() bool {
	return isCharDevice(os.Stdin) && isCharDevice(os.Stdout)
}

func isCharDevice(f *os.File) bool {
	info, err := f.Stat()
	if err != nil {
		return false
	}
	return info.Mode()&os.ModeCharDevice != 0
}

// agentMode reports whether a command should behave non-interactively: emit
// machine-readable output and never open a browser or block waiting on a
// human. True when --json is set or when no terminal is attached.
func agentMode(jsonFlag bool) bool {
	return jsonFlag || !isInteractive()
}

// pollUntil invokes check every interval until it returns done=true, the
// deadline passes, or ctx is cancelled. It waits one interval before the
// first check (matching the connect flow's "open browser, then poll" timing).
// On timeout it returns the error produced by timeoutMsg (evaluated lazily so
// callers can fold in the last transient error they observed).
func pollUntil(ctx context.Context, interval, timeout time.Duration, timeoutMsg func() string, check func() (bool, error)) error {
	deadline := time.Now().Add(timeout)
	for {
		select {
		case <-ctx.Done():
			return fmt.Errorf("interrupted")
		case <-time.After(interval):
		}

		if time.Now().After(deadline) {
			return fmt.Errorf("%s", timeoutMsg())
		}

		done, err := check()
		if err != nil {
			return err
		}
		if done {
			return nil
		}
	}
}

// timeoutMessage builds a pollUntil timeout message that appends the last
// transient error (if any) so a full-window outage is diagnosable rather than
// looking like the human never finished the browser step.
func timeoutMessage(base string, lastErr *error) func() string {
	return func() string {
		if lastErr != nil && *lastErr != nil {
			return fmt.Sprintf("%s (last error: %v)", base, *lastErr)
		}
		return base
	}
}
