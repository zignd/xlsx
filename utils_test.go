package xlsx

import (
	"errors"
	"testing"
)

func TestNumericToLetters(t *testing.T) {
	t.Parallel()

	tests := []struct {
		column int
		want   string
	}{
		{-1, ""},
		{0, "A"},
		{1, "B"},
		{25, "Z"},
		{26, "AA"},
		{27, "AB"},
		{51, "AZ"},
		{52, "BA"},
		{701, "ZZ"},
		{702, "AAA"},
		{16383, "XFD"},
	}

	for _, tt := range tests {
		if got := numericToLetters(tt.column); got != tt.want {
			t.Errorf("numericToLetters(%d) = %q, want %q", tt.column, got, tt.want)
		}
	}
}

func TestCreateIdentifierFromCoords(t *testing.T) {
	t.Parallel()

	tests := []struct {
		column int
		row    int
		want   string
	}{
		{0, 0, "A1"},
		{1, 1, "B2"},
		{25, 9, "Z10"},
		{26, 99, "AA100"},
		{16383, 1048575, "XFD1048576"},
	}

	for _, tt := range tests {
		if got := createIdentifierFromCoords(tt.column, tt.row); got != tt.want {
			t.Errorf("createIdentifierFromCoords(%d, %d) = %q, want %q", tt.column, tt.row, got, tt.want)
		}
	}
}

func TestXMLEscape(t *testing.T) {
	t.Parallel()

	got := xmlEscape(`a & b < c > d "e" 'f'`)
	want := "a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;"
	if got != want {
		t.Fatalf("xmlEscape() = %q, want %q", got, want)
	}

	if got := xmlEscape("héllo 世界"); got != "héllo 世界" {
		t.Fatalf("xmlEscape() changed UTF-8 text: %q", got)
	}

	if got := xmlEscape("a\x00b\x01c"); got != "abc" {
		t.Fatalf("xmlEscape() did not remove invalid XML characters: %q", got)
	}
}

func FuzzNumericToLetters(f *testing.F) {
	f.Add(0)
	f.Add(25)
	f.Add(26)
	f.Add(16383)

	f.Fuzz(func(t *testing.T, column int) {
		if column < 0 {
			return
		}
		letters := numericToLetters(column)
		if letters == "" {
			t.Fatalf("numericToLetters(%d) returned an empty string", column)
		}
		for _, r := range letters {
			if r < 'A' || r > 'Z' {
				t.Fatalf("numericToLetters(%d) = %q contains a non A-Z rune", column, letters)
			}
		}
	})
}

func TestParseCellReference(t *testing.T) {
	t.Parallel()

	tests := []struct {
		ref    string
		column int
		row    int
		ok     bool
	}{
		{"A1", 0, 1, true},
		{"$B$2", 1, 2, true},
		{"XFD1048576", 16383, 1_048_576, true},
		{"", 0, 0, false},
		{"A", 0, 0, false},
		{"1", 0, 0, false},
		{"A0", 0, 0, false},
		{"ABCDEFGH1", 0, 0, false},
	}
	for _, tt := range tests {
		column, row, err := parseCellReference(tt.ref)
		if tt.ok {
			if err != nil || column != tt.column || row != tt.row {
				t.Errorf("parseCellReference(%q) = (%d, %d, %v), want (%d, %d, nil)", tt.ref, column, row, err, tt.column, tt.row)
			}
		} else if err == nil {
			t.Errorf("parseCellReference(%q) = nil error, want error", tt.ref)
		}
	}
}

func TestNormalizeRange(t *testing.T) {
	t.Parallel()

	normalized, rect, err := normalizeRange("$B$2:C3", maxWorksheetRows, maxWorksheetColumns)
	if err != nil {
		t.Fatal(err)
	}
	if normalized != "B2:C3" {
		t.Fatalf("normalized = %q, want B2:C3", normalized)
	}
	if rect.minCol != 1 || rect.minRow != 2 || rect.maxCol != 2 || rect.maxRow != 3 {
		t.Fatalf("unexpected rect: %+v", rect)
	}
	for _, ref := range []string{"", "A1", "C3:B2", "A1:B2:C3", "A1:XFE"} {
		if _, _, err := normalizeRange(ref, maxWorksheetRows, maxWorksheetColumns); err != nil && ref == "A1" {
			t.Errorf("normalizeRange(%q) returned an unexpected error: %v", ref, err)
		}
	}
	if _, _, err := normalizeRange("C3:B2", maxWorksheetRows, maxWorksheetColumns); !errors.Is(err, ErrInvalidRange) {
		t.Fatalf("reversed range error = %v, want ErrInvalidRange", err)
	}
	if _, _, err := normalizeRange("A1:XFE", maxWorksheetRows, maxWorksheetColumns); !errors.Is(err, ErrInvalidRange) {
		t.Fatalf("out-of-range error = %v, want ErrInvalidRange", err)
	}
}
