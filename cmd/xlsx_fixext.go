package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// excelFormat represents the detected binary format of an Excel file.
type excelFormat int

const (
	excelFormatUnknown excelFormat = iota
	excelFormatOLE2                // Binary .xls (magic: d0cf11e0a1b11ae1)
	excelFormatOOXML               // ZIP-based .xlsx (magic: 504b0304)
)

// detectExcelFormat reads the first bytes of a file and returns the detected format.
func detectExcelFormat(filePath string) (excelFormat, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return excelFormatUnknown, err
	}
	defer f.Close()

	buf := make([]byte, 8)
	n, err := f.Read(buf)
	if err != nil {
		return excelFormatUnknown, err
	}
	if n < 4 {
		return excelFormatUnknown, nil
	}

	// OLE2 Compound Document: d0 cf 11 e0 (full signature: d0cf11e0a1b11ae1)
	if buf[0] == 0xd0 && buf[1] == 0xcf && buf[2] == 0x11 && buf[3] == 0xe0 {
		return excelFormatOLE2, nil
	}

	// ZIP (OOXML): PK\x03\x04
	if buf[0] == 0x50 && buf[1] == 0x4b && buf[2] == 0x03 && buf[3] == 0x04 {
		return excelFormatOOXML, nil
	}

	return excelFormatUnknown, nil
}

// fixExcelExtension checks whether a file's extension matches its actual content.
// If there is a mismatch (.xls with OOXML content or .xlsx with OLE2 content),
// it renames the file on disk and returns the new path. A note is emitted to stderr.
// If the extension matches or the file is not .xls/.xlsx, it returns the path unchanged.
func fixExcelExtension(filePath string) (string, error) {
	ext := strings.ToLower(filepath.Ext(filePath))
	if ext != ".xls" && ext != ".xlsx" {
		return filePath, nil
	}

	format, err := detectExcelFormat(filePath)
	if err != nil {
		return filePath, err
	}

	if format == excelFormatUnknown {
		return filePath, nil
	}

	var newPath string
	switch {
	case ext == ".xls" && format == excelFormatOOXML:
		newPath = filePath + "x" // .xls → .xlsx
	case ext == ".xlsx" && format == excelFormatOLE2:
		newPath = strings.TrimSuffix(filePath, filepath.Ext(filePath)) + ".xls" // .xlsx → .xls
	default:
		return filePath, nil // extension matches content
	}

	// Don't silently overwrite an existing file
	if _, err := os.Stat(newPath); err == nil {
		return "", fmt.Errorf("cannot rename %s to %s: target already exists", filepath.Base(filePath), filepath.Base(newPath))
	}

	if err := os.Rename(filePath, newPath); err != nil {
		return "", fmt.Errorf("renaming %s: %w", filepath.Base(filePath), err)
	}

	formatName := "OOXML"
	if format == excelFormatOLE2 {
		formatName = "OLE2"
	}
	fmt.Fprintf(os.Stderr, "note: %s is %s format — renamed to %s\n", filepath.Base(filePath), formatName, filepath.Base(newPath))

	return newPath, nil
}

// fixWritebackExtension checks a file that was just written back by the server.
// If the server converted OLE2→OOXML, the written bytes
// may not match the file extension. This renames to match.
func fixWritebackExtension(filePath string) (string, error) {
	ext := strings.ToLower(filepath.Ext(filePath))
	if ext != ".xls" && ext != ".xlsx" {
		return filePath, nil
	}

	format, err := detectExcelFormat(filePath)
	if err != nil {
		return filePath, err
	}

	if format == excelFormatUnknown {
		return filePath, nil
	}

	var newPath string
	switch {
	case ext == ".xls" && format == excelFormatOOXML:
		newPath = filePath + "x"
	case ext == ".xlsx" && format == excelFormatOLE2:
		newPath = strings.TrimSuffix(filePath, filepath.Ext(filePath)) + ".xls"
	default:
		return filePath, nil
	}

	if _, err := os.Stat(newPath); err == nil {
		return "", fmt.Errorf("cannot rename %s to %s: target already exists", filepath.Base(filePath), filepath.Base(newPath))
	}

	if err := os.Rename(filePath, newPath); err != nil {
		return "", fmt.Errorf("renaming %s: %w", filepath.Base(filePath), err)
	}

	fmt.Fprintf(os.Stderr, "note: converted output saved as %s\n", filepath.Base(newPath))

	return newPath, nil
}
