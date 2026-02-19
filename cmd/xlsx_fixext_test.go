package cmd

import (
	"os"
	"path/filepath"
	"testing"
)

func TestDetectExcelFormat(t *testing.T) {
	tests := []struct {
		name   string
		header []byte
		want   excelFormat
	}{
		{
			name:   "OLE2 magic bytes",
			header: []byte{0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1},
			want:   excelFormatOLE2,
		},
		{
			name:   "ZIP/OOXML magic bytes",
			header: []byte{0x50, 0x4b, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00},
			want:   excelFormatOOXML,
		},
		{
			name:   "unknown format",
			header: []byte{0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07},
			want:   excelFormatUnknown,
		},
		{
			name:   "too short",
			header: []byte{0xd0, 0xcf},
			want:   excelFormatUnknown,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := filepath.Join(t.TempDir(), "test.bin")
			if err := os.WriteFile(f, tt.header, 0o644); err != nil {
				t.Fatal(err)
			}
			got, err := detectExcelFormat(f)
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if got != tt.want {
				t.Errorf("detectExcelFormat = %d, want %d", got, tt.want)
			}
		})
	}
}

func TestFixExcelExtension(t *testing.T) {
	ole2Header := []byte{0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1}
	ooxmlHeader := []byte{0x50, 0x4b, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00}

	t.Run("xls with OOXML content renames to xlsx", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "budget.xls")
		if err := os.WriteFile(f, ooxmlHeader, 0o644); err != nil {
			t.Fatal(err)
		}

		got, err := fixExcelExtension(f)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}

		want := filepath.Join(dir, "budget.xlsx")
		if got != want {
			t.Errorf("got %q, want %q", got, want)
		}
		if _, err := os.Stat(want); err != nil {
			t.Errorf("renamed file does not exist: %v", err)
		}
		if _, err := os.Stat(f); !os.IsNotExist(err) {
			t.Errorf("original file still exists")
		}
	})

	t.Run("xlsx with OLE2 content renames to xls", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "data.xlsx")
		if err := os.WriteFile(f, ole2Header, 0o644); err != nil {
			t.Fatal(err)
		}

		got, err := fixExcelExtension(f)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}

		want := filepath.Join(dir, "data.xls")
		if got != want {
			t.Errorf("got %q, want %q", got, want)
		}
		if _, err := os.Stat(want); err != nil {
			t.Errorf("renamed file does not exist: %v", err)
		}
	})

	t.Run("xls with OLE2 content is no-op", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "correct.xls")
		if err := os.WriteFile(f, ole2Header, 0o644); err != nil {
			t.Fatal(err)
		}

		got, err := fixExcelExtension(f)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got != f {
			t.Errorf("got %q, want %q (should be unchanged)", got, f)
		}
	})

	t.Run("xlsx with OOXML content is no-op", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "correct.xlsx")
		if err := os.WriteFile(f, ooxmlHeader, 0o644); err != nil {
			t.Fatal(err)
		}

		got, err := fixExcelExtension(f)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got != f {
			t.Errorf("got %q, want %q (should be unchanged)", got, f)
		}
	})

	t.Run("non-Excel extension is no-op", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "data.csv")
		if err := os.WriteFile(f, ooxmlHeader, 0o644); err != nil {
			t.Fatal(err)
		}

		got, err := fixExcelExtension(f)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got != f {
			t.Errorf("got %q, want %q (should be unchanged)", got, f)
		}
	})

	t.Run("errors if target already exists", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "budget.xls")
		if err := os.WriteFile(f, ooxmlHeader, 0o644); err != nil {
			t.Fatal(err)
		}
		// Create the target file so rename would collide
		target := filepath.Join(dir, "budget.xlsx")
		if err := os.WriteFile(target, []byte("existing"), 0o644); err != nil {
			t.Fatal(err)
		}

		_, err := fixExcelExtension(f)
		if err == nil {
			t.Fatal("expected error when target exists, got nil")
		}
	})
}

func TestFixWritebackExtension(t *testing.T) {
	ooxmlHeader := []byte{0x50, 0x4b, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00}

	t.Run("xls with OOXML writeback renames to xlsx", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "budget.xls")
		if err := os.WriteFile(f, ooxmlHeader, 0o644); err != nil {
			t.Fatal(err)
		}

		got, err := fixWritebackExtension(f)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}

		want := filepath.Join(dir, "budget.xlsx")
		if got != want {
			t.Errorf("got %q, want %q", got, want)
		}
		if _, err := os.Stat(want); err != nil {
			t.Errorf("renamed file does not exist: %v", err)
		}
	})

	t.Run("xlsx with OOXML writeback is no-op", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "correct.xlsx")
		if err := os.WriteFile(f, ooxmlHeader, 0o644); err != nil {
			t.Fatal(err)
		}

		got, err := fixWritebackExtension(f)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got != f {
			t.Errorf("got %q, want %q (should be unchanged)", got, f)
		}
	})

	t.Run("errors if target already exists", func(t *testing.T) {
		dir := t.TempDir()
		f := filepath.Join(dir, "budget.xls")
		if err := os.WriteFile(f, ooxmlHeader, 0o644); err != nil {
			t.Fatal(err)
		}
		target := filepath.Join(dir, "budget.xlsx")
		if err := os.WriteFile(target, []byte("existing"), 0o644); err != nil {
			t.Fatal(err)
		}

		_, err := fixWritebackExtension(f)
		if err == nil {
			t.Fatal("expected error when target exists, got nil")
		}
	})
}
