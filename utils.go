package xlsx

import (
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
