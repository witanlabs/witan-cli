package cmd

import (
	"encoding/json"
	"testing"

	"github.com/witanlabs/witan-cli/client"
)

func TestParseEditCell(t *testing.T) {
	tests := []struct {
		name         string
		arg          string
		globalFormat string
		want         client.EditCell
		wantErr      bool
	}{
		{
			name: "number value",
			arg:  "Sheet1!A1=42",
			want: client.EditCell{Address: "Sheet1!A1", Value: json.RawMessage("42")},
		},
		{
			name: "float value",
			arg:  "Sheet1!B2=3.14",
			want: client.EditCell{Address: "Sheet1!B2", Value: json.RawMessage("3.14")},
		},
		{
			name: "formula via double equals",
			arg:  "Sheet1!A1==SUM(A:A)",
			want: client.EditCell{Address: "Sheet1!A1", Formula: "=SUM(A:A)"},
		},
		{
			name: "string value",
			arg:  "Sheet1!A1=hello",
			want: client.EditCell{Address: "Sheet1!A1", Value: json.RawMessage(`"hello"`)},
		},
		{
			name: "boolean true",
			arg:  "Sheet1!A1=true",
			want: client.EditCell{Address: "Sheet1!A1", Value: json.RawMessage("true")},
		},
		{
			name: "boolean TRUE (case insensitive)",
			arg:  "Sheet1!A1=TRUE",
			want: client.EditCell{Address: "Sheet1!A1", Value: json.RawMessage("true")},
		},
		{
			name: "boolean false",
			arg:  "Sheet1!A1=false",
			want: client.EditCell{Address: "Sheet1!A1", Value: json.RawMessage("false")},
		},
		{
			name: "null clears cell",
			arg:  "Sheet1!A1=null",
			want: client.EditCell{Address: "Sheet1!A1", Value: json.RawMessage("null")},
		},
		{
			name:         "format-only bare address",
			arg:          "A1",
			globalFormat: "#,##0",
			want:         client.EditCell{Address: "A1", Format: "#,##0"},
		},
		{
			name:         "value with global format",
			arg:          "Sheet1!A1=42",
			globalFormat: "#,##0.00",
			want:         client.EditCell{Address: "Sheet1!A1", Value: json.RawMessage("42"), Format: "#,##0.00"},
		},
		{
			name:         "formula with global format",
			arg:          "Sheet1!A1==SUM(B1:B10)",
			globalFormat: "0.00%",
			want:         client.EditCell{Address: "Sheet1!A1", Formula: "=SUM(B1:B10)", Format: "0.00%"},
		},
		{
			name: "sheet name with equals sign",
			arg:  "My=Sheet!A1=42",
			want: client.EditCell{Address: "My=Sheet!A1", Value: json.RawMessage("42")},
		},
		{
			name: "no sheet separator parses address as-is",
			arg:  "A1=42",
			want: client.EditCell{Address: "A1", Value: json.RawMessage("42")},
		},
		{
			name:    "empty address",
			arg:     "=42",
			wantErr: true,
		},
		{
			name:    "bare address without format errors",
			arg:     "Sheet1!A1",
			wantErr: true,
		},
		{
			name:    "completely empty arg without format",
			arg:     "",
			wantErr: true,
		},
		{
			name: "negative number",
			arg:  "Sheet1!A1=-5",
			want: client.EditCell{Address: "Sheet1!A1", Value: json.RawMessage("-5")},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := parseEditCell(tt.arg, tt.globalFormat)
			if tt.wantErr {
				if err == nil {
					t.Fatalf("expected error, got nil")
				}
				return
			}
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if got.Address != tt.want.Address {
				t.Errorf("Address = %q, want %q", got.Address, tt.want.Address)
			}
			if got.Formula != tt.want.Formula {
				t.Errorf("Formula = %q, want %q", got.Formula, tt.want.Formula)
			}
			if got.Format != tt.want.Format {
				t.Errorf("Format = %q, want %q", got.Format, tt.want.Format)
			}
			if string(got.Value) != string(tt.want.Value) {
				t.Errorf("Value = %s, want %s", string(got.Value), string(tt.want.Value))
			}
		})
	}
}
