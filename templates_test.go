package xlsx

import (
	"testing"
	"time"
)

func TestTimeToExcelTime(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name string
		time time.Time
		want float64
	}{
		{"unix epoch", time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC), 25569},
		{"midnight", time.Date(2024, 1, 2, 0, 0, 0, 0, time.UTC), 45293},
		{"noon", time.Date(1970, 1, 1, 12, 0, 0, 0, time.UTC), 25569.5},
		{"fractional second", time.Date(1970, 1, 1, 0, 0, 0, 500000000, time.UTC), 25569 + 0.5/86400},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			if got := timeToExcelTime(tt.time); got != tt.want {
				t.Errorf("timeToExcelTime(%s) = %v, want %v", tt.time, got, tt.want)
			}
		})
	}
}

func TestTimeToExcelTimeUsesWallClockComponents(t *testing.T) {
	t.Parallel()

	utc := time.Date(2024, 6, 15, 10, 30, 0, 0, time.UTC)
	withOffset := time.Date(2024, 6, 15, 10, 30, 0, 0, time.FixedZone("UTC+5", 5*60*60))

	if got, want := timeToExcelTime(withOffset), timeToExcelTime(utc); got != want {
		t.Fatalf("timeToExcelTime() depends on the time zone: got %v, want %v", got, want)
	}
}

func TestCellFormats(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name string
		got  string
		want string
	}{
		{
			name: "text escapes XML",
			got:  textCellXML("A1", `a & b <c>`, 0),
			want: `<c r="A1" t="inlineStr"><is><t xml:space="preserve">a &amp; b &lt;c&gt;</t></is></c>`,
		},
		{
			name: "number",
			got:  numberCellXML("B2", "42.5", 0),
			want: `<c r="B2" t="n"><v>42.5</v></c>`,
		},
		{
			name: "bool true",
			got:  boolCellXML("C3", true, 0),
			want: `<c r="C3" t="b"><v>1</v></c>`,
		},
		{
			name: "bool false",
			got:  boolCellXML("C3", false, 0),
			want: `<c r="C3" t="b"><v>0</v></c>`,
		},
		{
			name: "blank",
			got:  blankCellXML("D4", 0),
			want: `<c r="D4"/>`,
		},
		{
			name: "date",
			got:  dateCellXML("E5", time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC), 0),
			want: `<c r="E5" s="1" t="n"><v>25569</v></c>`,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			if tt.got != tt.want {
				t.Errorf("got %q, want %q", tt.got, tt.want)
			}
		})
	}
}

func TestSheetAndRelationshipFormats(t *testing.T) {
	t.Parallel()

	if got, want := sheetXML(`A&B`, 2), `<sheet name="A&amp;B" sheetId="2" r:id="rId2"/>`; got != want {
		t.Errorf("sheetXML() = %q, want %q", got, want)
	}

	got := relationshipXML(3, "https://example.com/type?a=1&b=2", "worksheets/sheet1.xml")
	want := `<Relationship Id="rId3" Type="https://example.com/type?a=1&amp;b=2" Target="worksheets/sheet1.xml"/>`
	if got != want {
		t.Errorf("relationshipXML() = %q, want %q", got, want)
	}
}
