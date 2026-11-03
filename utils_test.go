package xlsx

import "testing"

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
