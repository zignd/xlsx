package xlsx

import (
	"fmt"
	"strconv"
	"strings"
)

// createIdentifierFromCoords returns an A1-style cell reference from
// zero-based column and row coordinates.
func createIdentifierFromCoords(x, y int) string {
	return numericToLetters(x) + strconv.Itoa(y+1)
}

// numericToLetters converts a zero-based column index into its spreadsheet
// letters (0 -> A, 25 -> Z, 26 -> AA, 701 -> ZZ, 702 -> AAA).
func numericToLetters(colRef int) string {
	if colRef < 0 {
		return ""
	}

	// Excel columns are bijective base-26: there is no zero digit.
	var reversed []byte
	for colRef >= 0 {
		reversed = append(reversed, byte(colRef%26)+'A')
		colRef = colRef/26 - 1
	}

	for i, j := 0, len(reversed)-1; i < j; i, j = i+1, j-1 {
		reversed[i], reversed[j] = reversed[j], reversed[i]
	}
	return string(reversed)
}

func xmlEscape(s string) string {
	var b strings.Builder
	b.Grow(len(s))
	for _, r := range s {
		switch r {
		case '&':
			b.WriteString("&amp;")
		case '<':
			b.WriteString("&lt;")
		case '>':
			b.WriteString("&gt;")
		case '"':
			b.WriteString("&quot;")
		case '\'':
			b.WriteString("&apos;")
		default:
			if validXMLRune(r) {
				b.WriteRune(r)
			}
		}
	}
	return b.String()
}

func validXMLRune(r rune) bool {
	return r == '\t' || r == '\n' || r == '\r' ||
		(r >= 0x20 && r <= 0xD7FF) ||
		(r >= 0xE000 && r <= 0xFFFD) ||
		(r >= 0x10000 && r <= 0x10FFFF)
}

type cellRect struct {
	minCol int
	minRow int
	maxCol int
	maxRow int
}

func parseCellReference(ref string) (column, row int, err error) {
	ref = strings.ToUpper(strings.TrimSpace(ref))
	ref = strings.ReplaceAll(ref, "$", "")
	if ref == "" {
		return 0, 0, fmt.Errorf("%w: empty reference", ErrInvalidCellReference)
	}

	i := 0
	for i < len(ref) && ref[i] >= 'A' && ref[i] <= 'Z' {
		i++
	}
	if i == 0 || i == len(ref) {
		return 0, 0, fmt.Errorf("%w: %q", ErrInvalidCellReference, ref)
	}

	column, err = lettersToNumeric(ref[:i])
	if err != nil {
		return 0, 0, err
	}
	row, err = strconv.Atoi(ref[i:])
	if err != nil || row < 1 {
		return 0, 0, fmt.Errorf("%w: %q", ErrInvalidCellReference, ref)
	}
	return column, row, nil
}

func lettersToNumeric(letters string) (int, error) {
	if letters == "" || len(letters) > 7 {
		return 0, fmt.Errorf("%w: %q", ErrInvalidCellReference, letters)
	}
	column := 0
	for _, r := range strings.ToUpper(letters) {
		if r < 'A' || r > 'Z' {
			return 0, fmt.Errorf("%w: %q", ErrInvalidCellReference, letters)
		}
		column = column*26 + int(r-'A'+1)
	}
	return column - 1, nil
}

func normalizeRange(ref string, maxRows, maxColumns int) (string, cellRect, error) {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return "", cellRect{}, fmt.Errorf("%w: empty range", ErrInvalidRange)
	}

	parts := strings.Split(ref, ":")
	if len(parts) > 2 {
		return "", cellRect{}, fmt.Errorf("%w: %q", ErrInvalidRange, ref)
	}

	startCol, startRow, err := parseCellReference(parts[0])
	if err != nil {
		return "", cellRect{}, fmt.Errorf("%w: %q", ErrInvalidRange, ref)
	}
	endCol, endRow := startCol, startRow
	if len(parts) == 2 {
		endCol, endRow, err = parseCellReference(parts[1])
		if err != nil {
			return "", cellRect{}, fmt.Errorf("%w: %q", ErrInvalidRange, ref)
		}
	}

	if startCol > endCol || startRow > endRow {
		return "", cellRect{}, fmt.Errorf("%w: %q is reversed", ErrInvalidRange, ref)
	}
	if endCol >= maxColumns || endRow > maxRows || startRow < 1 {
		return "", cellRect{}, fmt.Errorf("%w: %q is outside the worksheet", ErrInvalidRange, ref)
	}

	normalized := createIdentifierFromCoords(startCol, startRow-1)
	if len(parts) == 2 {
		normalized += ":" + createIdentifierFromCoords(endCol, endRow-1)
	}
	return normalized, cellRect{
		minCol: startCol,
		minRow: startRow,
		maxCol: endCol,
		maxRow: endRow,
	}, nil
}

func rangesOverlap(a, b cellRect) bool {
	return a.minCol <= b.maxCol && b.minCol <= a.maxCol &&
		a.minRow <= b.maxRow && b.minRow <= a.maxRow
}
