package cmd

import (
	"strings"
	"testing"
)

func TestValidateSheetsExecArgs(t *testing.T) {
	tests := []struct {
		name    string
		create  bool
		args    []string
		wantErr string
	}{
		{
			name: "existing spreadsheet",
			args: []string{"gs://abc123"},
		},
		{
			name: "missing spreadsheet",
			args: []string{},
			wantErr: "requires exactly 1 spreadsheet reference",
		},
		{
			name:   "create without arg",
			create: true,
		},
		{
			name:   "create with new ref",
			create: true,
			args:   []string{"new"},
		},
		{
			name:   "create with gs new ref",
			create: true,
			args:   []string{"gs://new"},
		},
		{
			name:    "create with real ref",
			create:  true,
			args:    []string{"gs://abc123"},
			wantErr: "--create requires spreadsheet reference 'new' or gs://new",
		},
		{
			name: "implicit create via new",
			args: []string{"new"},
		},
		{
			name:    "too many args on create",
			create:  true,
			args:    []string{"new", "extra"},
			wantErr: "accepts at most 1 spreadsheet reference",
		},
		{
			name:    "invalid ref",
			args:    []string{"not-a-sheet"},
			wantErr: "invalid spreadsheet reference",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sheetsExecCreate = tt.create
			t.Cleanup(func() { sheetsExecCreate = false })

			err := validateSheetsExecArgs(sheetsExecCmd, tt.args)
			if tt.wantErr == "" {
				if err != nil {
					t.Fatalf("unexpected error: %v", err)
				}
				return
			}
			if err == nil {
				t.Fatalf("expected error containing %q", tt.wantErr)
			}
			if !strings.Contains(err.Error(), tt.wantErr) {
				t.Fatalf("error = %q, want substring %q", err.Error(), tt.wantErr)
			}
		})
	}
}

func TestResolveSheetsCreate(t *testing.T) {
	tests := []struct {
		name   string
		create bool
		args   []string
		want   bool
	}{
		{name: "flag", create: true, want: true},
		{name: "new ref", args: []string{"new"}, want: true},
		{name: "gs new ref", args: []string{"gs://new"}, want: true},
		{name: "existing ref", args: []string{"gs://abc"}, want: false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := resolveSheetsCreate(tt.create, tt.args); got != tt.want {
				t.Fatalf("resolveSheetsCreate() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestValidateSheetsTitle(t *testing.T) {
	if err := validateSheetsTitle(strings.Repeat("a", 1001), true); err == nil {
		t.Fatal("expected title length error")
	}
	if err := validateSheetsTitle("Budget", true); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestIsSheetsCreateRef(t *testing.T) {
	if !isSheetsCreateRef("new") {
		t.Fatal("expected new to be create ref")
	}
	if !isSheetsCreateRef("gs://new") {
		t.Fatal("expected gs://new to be create ref")
	}
	if isSheetsCreateRef("gs://abc123") {
		t.Fatal("expected real id not to be create ref")
	}
}
