package cmd

import (
	"context"
	"errors"
	"testing"
	"time"
)

func TestAgentModeJSONForcesAgent(t *testing.T) {
	// --json always means agent mode regardless of terminal state.
	if !agentMode(true) {
		t.Fatal("agentMode(true) = false, want true")
	}
}

func constMsg(s string) func() string {
	return func() string { return s }
}

func TestPollUntilDone(t *testing.T) {
	calls := 0
	err := pollUntil(context.Background(), time.Millisecond, time.Second, constMsg("timed out"), func() (bool, error) {
		calls++
		return calls >= 3, nil
	})
	if err != nil {
		t.Fatalf("pollUntil returned error: %v", err)
	}
	if calls != 3 {
		t.Fatalf("check called %d times, want 3", calls)
	}
}

func TestPollUntilTimeout(t *testing.T) {
	err := pollUntil(context.Background(), time.Millisecond, 5*time.Millisecond, constMsg("timed out waiting"), func() (bool, error) {
		return false, nil
	})
	if err == nil || err.Error() != "timed out waiting" {
		t.Fatalf("pollUntil error = %v, want \"timed out waiting\"", err)
	}
}

func TestPollUntilTimeoutIncludesLastError(t *testing.T) {
	var lastErr error
	err := pollUntil(context.Background(), time.Millisecond, 5*time.Millisecond,
		timeoutMessage("timed out waiting", &lastErr),
		func() (bool, error) {
			lastErr = errors.New("connection refused")
			return false, nil
		})
	if err == nil || err.Error() != "timed out waiting (last error: connection refused)" {
		t.Fatalf("pollUntil error = %v, want last-error suffix", err)
	}
}

func TestPollUntilCheckError(t *testing.T) {
	sentinel := errors.New("boom")
	err := pollUntil(context.Background(), time.Millisecond, time.Second, constMsg("timed out"), func() (bool, error) {
		return false, sentinel
	})
	if !errors.Is(err, sentinel) {
		t.Fatalf("pollUntil error = %v, want %v", err, sentinel)
	}
}

func TestPollUntilContextCancel(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	err := pollUntil(ctx, time.Millisecond, time.Second, constMsg("timed out"), func() (bool, error) {
		return false, nil
	})
	if err == nil || err.Error() != "interrupted" {
		t.Fatalf("pollUntil error = %v, want \"interrupted\"", err)
	}
}
