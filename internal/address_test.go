package internal

import (
	"testing"
)

func TestParseRange(t *testing.T) {
	tests := []struct {
		input                              string
		sheet                              string
		startRow, startCol, endRow, endCol int
		wantErr                            bool
	}{
		{"Sheet1!A1:Z50", "Sheet1", 1, 1, 50, 26, false},
		{"Sheet1!A1:B2", "Sheet1", 1, 1, 2, 2, false},
		{"Sheet1!A1", "Sheet1", 1, 1, 1, 1, false},
		{"'My Sheet'!C3:D4", "My Sheet", 3, 3, 4, 4, false},
		{"Sheet1!$A$1:$B$2", "Sheet1", 1, 1, 2, 2, false},
		// reversed range should normalize
		{"Sheet1!B2:A1", "Sheet1", 1, 1, 2, 2, false},
		// missing sheet
		{"A1:B2", "", 0, 0, 0, 0, true},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			sheet, sr, sc, er, ec, err := ParseRange(tt.input)
			if tt.wantErr {
				if err == nil {
					t.Fatalf("expected error for %q", tt.input)
				}
				return
			}
			if err != nil {
				t.Fatalf("unexpected error for %q: %v", tt.input, err)
			}
			if sheet != tt.sheet || sr != tt.startRow || sc != tt.startCol || er != tt.endRow || ec != tt.endCol {
				t.Errorf("ParseRange(%q) = (%q, %d, %d, %d, %d), want (%q, %d, %d, %d, %d)",
					tt.input, sheet, sr, sc, er, ec,
					tt.sheet, tt.startRow, tt.startCol, tt.endRow, tt.endCol)
			}
		})
	}
}

func TestColToLetter(t *testing.T) {
	tests := []struct {
		col  int
		want string
	}{
		{1, "A"},
		{26, "Z"},
		{27, "AA"},
		{52, "AZ"},
		{702, "ZZ"},
	}
	for _, tt := range tests {
		if got := ColToLetter(tt.col); got != tt.want {
			t.Errorf("ColToLetter(%d) = %q, want %q", tt.col, got, tt.want)
		}
	}
}

func TestFormatAddress(t *testing.T) {
	got := FormatAddress("Sheet1", 1, 1, 50, 26)
	want := "Sheet1!A1:Z50"
	if got != want {
		t.Errorf("FormatAddress = %q, want %q", got, want)
	}

	// Single cell
	got = FormatAddress("Sheet1", 5, 3, 5, 3)
	want = "Sheet1!C5"
	if got != want {
		t.Errorf("FormatAddress single cell = %q, want %q", got, want)
	}
}
