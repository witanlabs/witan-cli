package internal

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// cellRefRe matches a cell reference like A1, $B$2, AA100
var cellRefRe = regexp.MustCompile(`^\$?([A-Z]+)\$?(\d+)$`)

// ParseRange parses an address like "Sheet1!A1:Z50" and returns
// (sheet, startRow, startCol, endRow, endCol) in 1-indexed form.
func ParseRange(address string) (sheet string, startRow, startCol, endRow, endCol int, err error) {
	// Split sheet!range
	sheetPart, rangePart, hasSheet := strings.Cut(address, "!")
	if !hasSheet {
		return "", 0, 0, 0, 0, fmt.Errorf("address must include sheet name (e.g. Sheet1!A1:B2), got %q", address)
	}

	// Remove surrounding quotes from sheet name
	sheet = strings.Trim(sheetPart, "'")

	// Split range into from:to
	fromRef, toRef, hasColon := strings.Cut(rangePart, ":")
	if !hasColon {
		toRef = fromRef // single cell
	}

	startCol, startRow, err = parseRef(fromRef)
	if err != nil {
		return "", 0, 0, 0, 0, fmt.Errorf("invalid start of range %q: %w", fromRef, err)
	}
	endCol, endRow, err = parseRef(toRef)
	if err != nil {
		return "", 0, 0, 0, 0, fmt.Errorf("invalid end of range %q: %w", toRef, err)
	}

	// Normalize order
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}

	return sheet, startRow, startCol, endRow, endCol, nil
}

// ColToLetter converts a 1-indexed column number to Excel letter(s)
func ColToLetter(col int) string {
	result := ""
	for col > 0 {
		col--
		result = string(rune('A'+col%26)) + result
		col /= 26
	}
	return result
}

// FormatAddress builds an address string like "Sheet1!A1:Z50"
func FormatAddress(sheet string, startRow, startCol, endRow, endCol int) string {
	from := ColToLetter(startCol) + strconv.Itoa(startRow)
	to := ColToLetter(endCol) + strconv.Itoa(endRow)
	if from == to {
		return sheet + "!" + from
	}
	return sheet + "!" + from + ":" + to
}

func parseRef(ref string) (col, row int, err error) {
	ref = strings.ReplaceAll(ref, "$", "")
	m := cellRefRe.FindStringSubmatch(strings.ToUpper(ref))
	if m == nil {
		return 0, 0, fmt.Errorf("invalid cell reference %q", ref)
	}
	col = letterToCol(m[1])
	row, _ = strconv.Atoi(m[2])
	return col, row, nil
}

func letterToCol(letters string) int {
	col := 0
	for _, c := range letters {
		col = col*26 + int(c-'A'+1)
	}
	return col
}
