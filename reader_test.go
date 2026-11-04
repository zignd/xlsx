package xlsx

import (
	"archive/zip"
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"sort"
	"testing"
	"time"
)

func writeRawXLSX(t *testing.T, parts map[string]string) string {
	t.Helper()

	path := filepath.Join(t.TempDir(), "raw.xlsx")
	file, err := os.Create(path)
	if err != nil {
		t.Fatalf("Create() error = %v", err)
	}
	writer := zip.NewWriter(file)
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		entry, err := writer.Create(name)
		if err != nil {
			t.Fatalf("zip Create(%q) error = %v", name, err)
		}
		if _, err := entry.Write([]byte(parts[name])); err != nil {
			t.Fatalf("zip Write(%q) error = %v", name, err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatalf("zip Close() error = %v", err)
	}
	if err := file.Close(); err != nil {
		t.Fatalf("file Close() error = %v", err)
	}
	return path
}

func minimalStylesXML() string {
	return xmlDeclaration + `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>` +
		`<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
		`<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
		`<borders count="1"><border/></borders>` +
		`<cellStyleXfs count="1"><xf numFmtId="0"/></cellStyleXfs>` +
		`<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs>` +
		`</styleSheet>`
}

func minimalPackage(sheetContent string) map[string]string {
	return map[string]string{
		"[Content_Types].xml": contentTypesStart +
			contentTypeOverride("/xl/workbook.xml", contentTypeWorkbook) +
			contentTypeOverride("/xl/worksheets/sheet1.xml", contentTypeWorksheet) +
			contentTypeOverride("/xl/styles.xml", contentTypeStyles) +
			contentTypesEnd,
		"_rels/.rels":     rootRelationships,
		"xl/workbook.xml": workbookStart + sheetXML("Data", 1) + workbookEnd,
		"xl/_rels/workbook.xml.rels": workbookRelsStart +
			relationshipXML(1, contentTypeWorksheetR, "worksheets/sheet1.xml") +
			relationshipXML(2, contentTypeStylesR, "styles.xml") +
			workbookRelsEnd,
		"xl/styles.xml": minimalStylesXML(),
		"xl/worksheets/sheet1.xml": worksheetStartPrefix +
			`<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
			worksheetDataStart + sheetContent + worksheetDataEnd + worksheetEnd,
	}
}

func writeFeatureWorkbook(t *testing.T) string {
	t.Helper()

	path := filepath.Join(t.TempDir(), "features.xlsx")
	workbook, err := NewWorkbookWithOptions(path, Options{SharedStrings: true})
	if err != nil {
		t.Fatalf("NewWorkbookWithOptions() error = %v", err)
	}
	styleID, err := workbook.AddStyle(Style{
		Bold:         true,
		NumberFormat: "#,##0.00",
		FillColor:    "FFF2CC",
		Horizontal:   "center",
	})
	if err != nil {
		t.Fatalf("AddStyle() error = %v", err)
	}

	data := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := data.FreezePanes(1, 0); err != nil {
		t.Fatalf("FreezePanes() error = %v", err)
	}
	if err := data.DefineColumns([]*WorksheetColumn{
		{Key: "name", Value: "Name", Width: 24},
		{Key: "amount", Value: "Amount"},
		{Key: "when", Value: "When"},
		{Key: "active", Value: "Active"},
	}); err != nil {
		t.Fatalf("DefineColumns() error = %v", err)
	}
	if err := data.SetAutoFilter("A1:D3"); err != nil {
		t.Fatalf("SetAutoFilter() error = %v", err)
	}

	when := time.Date(2024, 1, 2, 3, 4, 5, 0, time.UTC)
	for i := 0; i < 2; i++ {
		row, err := data.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		if err := row.SetHeight(18); err != nil {
			t.Fatalf("SetHeight() error = %v", err)
		}
		name, _ := row.AddCellWithKey("name")
		name.Value = "Widget & <" + string(rune('1'+i)) + ">"
		amount, _ := row.AddCellWithKey("amount")
		amount.Value = 1.25 + float64(i)
		amount.SetStyle(styleID)
		dateCell, _ := row.AddCellWithKey("when")
		dateCell.Value = when.AddDate(0, 0, i)
		active, _ := row.AddCellWithKey("active")
		active.Value = i == 0
	}
	if err := data.CommitRows(); err != nil {
		t.Fatalf("data CommitRows() error = %v", err)
	}

	summary := workbook.AddWorksheet(&WorksheetOptions{Name: "Summary"})
	row, err := summary.AddRow()
	if err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	if err := row.Append("total", 42, 3.5, true, when, nil); err != nil {
		t.Fatalf("Append() error = %v", err)
	}
	formula, _ := row.AddCell()
	formula.SetFormula("=SUM(B1:B2)")
	if err := summary.CommitRows(); err != nil {
		t.Fatalf("summary CommitRows() error = %v", err)
	}

	if err := workbook.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	return path
}

func TestReaderGeneratedWorkbook(t *testing.T) {
	t.Parallel()

	path := writeFeatureWorkbook(t)
	if err := Validate(path); err != nil {
		t.Fatalf("Validate() error = %v", err)
	}

	reader, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer reader.Close()

	sheets := reader.Sheets()
	if got, want := len(sheets), 2; got != want {
		t.Fatalf("Sheets() length = %d, want %d", got, want)
	}
	if sheets[0].Name != "Data" || sheets[1].Name != "Summary" {
		t.Fatalf("unexpected sheets: %+v", sheets)
	}

	data, err := reader.OpenSheet("Data")
	if err != nil {
		t.Fatalf("OpenSheet(Data) error = %v", err)
	}
	defer data.Close()

	if !data.Next() {
		t.Fatalf("first Next() = false, err = %v", data.Err())
	}
	header := data.Row()
	if header.Index != 1 {
		t.Fatalf("header row = %d, want 1", header.Index)
	}
	headerName, ok := header.Cell("A1")
	if !ok || headerName.Type != CellTypeString {
		t.Fatalf("A1 = %+v, want string", headerName)
	}
	if value, _ := headerName.String(); value != "Name" {
		t.Fatalf("A1 = %q, want Name", value)
	}

	if !data.Next() {
		t.Fatalf("second Next() = false, err = %v", data.Err())
	}
	row := data.Row()
	if row.Index != 2 {
		t.Fatalf("data row = %d, want 2", row.Index)
	}
	name, ok := row.Cell("A2")
	if !ok {
		t.Fatal("A2 is missing")
	}
	if value, _ := name.String(); value != "Widget & <1>" {
		t.Fatalf("A2 = %q, want Widget & <1>", value)
	}
	amount, ok := row.Cell("B2")
	if !ok || amount.Type != CellTypeNumber {
		t.Fatalf("B2 = %+v, want number", amount)
	}
	if value, _ := amount.Float64(); value != 1.25 {
		t.Fatalf("B2 = %v, want 1.25", value)
	}
	if amount.StyleID == 0 {
		t.Fatal("B2 has no style")
	}
	when, ok := row.Cell("C2")
	if !ok || when.Type != CellTypeDate {
		t.Fatalf("C2 = %+v, want date", when)
	}
	whenValue, _ := when.Time()
	if !whenValue.Equal(time.Date(2024, 1, 2, 3, 4, 5, 0, time.UTC)) {
		t.Fatalf("C2 = %v, want 2024-01-02 03:04:05 UTC", whenValue)
	}
	active, ok := row.Cell("D2")
	if !ok || active.Type != CellTypeBool {
		t.Fatalf("D2 = %+v, want bool", active)
	}
	if value, _ := active.Bool(); !value {
		t.Fatal("D2 = false, want true")
	}

	if !data.Next() {
		t.Fatalf("third Next() = false, err = %v", data.Err())
	}
	if data.Row().Index != 3 {
		t.Fatalf("third row = %d, want 3", data.Row().Index)
	}
	if data.Next() {
		t.Fatal("unexpected fourth row")
	}
	if err := data.Err(); err != nil {
		t.Fatalf("Data Err() = %v", err)
	}

	summary, err := reader.OpenSheet("Summary")
	if err != nil {
		t.Fatalf("OpenSheet(Summary) error = %v", err)
	}
	defer summary.Close()
	if !summary.Next() {
		t.Fatalf("summary Next() = false, err = %v", summary.Err())
	}
	summaryRow := summary.Row()
	for ref, want := range map[string]string{
		"A1": "total",
		"D1": "true",
	} {
		cell, ok := summaryRow.Cell(ref)
		if !ok {
			t.Fatalf("%s is missing", ref)
		}
		if got, _ := cell.String(); got != want {
			t.Fatalf("%s = %q, want %q", ref, got, want)
		}
	}
	formula, ok := summaryRow.Cell("G1")
	if !ok || formula.Formula != "SUM(B1:B2)" {
		t.Fatalf("G1 = %+v, want formula SUM(B1:B2)", formula)
	}
}

func TestReaderSharedStringsModes(t *testing.T) {
	t.Parallel()

	path := writeFeatureWorkbook(t)
	for _, mode := range []SharedStringsMode{SharedStringsMemory, SharedStringsDisk, SharedStringsAuto} {
		mode := mode
		t.Run([]string{"memory", "disk", "auto"}[mode], func(t *testing.T) {
			t.Parallel()
			reader, err := Open(path, WithSharedStringsMode(mode))
			if err != nil {
				t.Fatalf("Open() error = %v", err)
			}
			defer reader.Close()
			sheet, err := reader.OpenSheet("Data")
			if err != nil {
				t.Fatalf("OpenSheet() error = %v", err)
			}
			defer sheet.Close()
			if !sheet.Next() {
				t.Fatalf("first Next() failed: %v", sheet.Err())
			}
			if !sheet.Next() {
				t.Fatalf("second Next() failed: %v", sheet.Err())
			}
			cell, ok := sheet.Row().Cell("A2")
			if !ok {
				t.Fatal("A2 is missing")
			}
			if value, _ := cell.String(); value != "Widget & <1>" {
				t.Fatalf("A2 = %q", value)
			}
		})
	}
}

func TestReaderHandWrittenSparseSheet(t *testing.T) {
	t.Parallel()

	parts := minimalPackage(`<row r="2"><c r="B2" t="s"><v>0</v></c><c r="C2"><v>42</v></c><c r="D2" s="1"><v>45293.5</v></c></row><row><c r="A3" t="inlineStr"><is><t>inline</t></is></c><c r="B3" t="b"><v>1</v></c></row>`)
	parts["xl/sharedStrings.xml"] = xmlDeclaration + `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>shared</t></si></sst>`
	parts["[Content_Types].xml"] = contentTypesStart +
		contentTypeOverride("/xl/workbook.xml", contentTypeWorkbook) +
		contentTypeOverride("/xl/worksheets/sheet1.xml", contentTypeWorksheet) +
		contentTypeOverride("/xl/styles.xml", contentTypeStyles) +
		contentTypeOverride("/xl/sharedStrings.xml", contentTypeSharedStrings) +
		contentTypesEnd
	parts["xl/_rels/workbook.xml.rels"] = workbookRelsStart +
		relationshipXML(1, contentTypeWorksheetR, "worksheets/sheet1.xml") +
		relationshipXML(2, contentTypeStylesR, "styles.xml") +
		relationshipXML(3, relationshipSharedStrings, "sharedStrings.xml") +
		workbookRelsEnd

	path := writeRawXLSX(t, parts)
	if err := Validate(path); err != nil {
		t.Fatalf("Validate() error = %v", err)
	}

	reader, err := Open(path, WithSharedStringsMode(SharedStringsMemory))
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer reader.Close()
	sheet, err := reader.OpenSheet("Data")
	if err != nil {
		t.Fatalf("OpenSheet() error = %v", err)
	}
	defer sheet.Close()

	if !sheet.Next() {
		t.Fatalf("Next() = false, err = %v", sheet.Err())
	}
	row := sheet.Row()
	if row.Index != 2 {
		t.Fatalf("row index = %d, want 2", row.Index)
	}
	shared, ok := row.Cell("B2")
	if !ok || shared.Type != CellTypeString {
		t.Fatalf("B2 = %+v", shared)
	}
	if value, _ := shared.String(); value != "shared" {
		t.Fatalf("B2 = %q", value)
	}
	number, ok := row.Cell("C2")
	if !ok || number.Type != CellTypeNumber {
		t.Fatalf("C2 = %+v", number)
	}
	if value, _ := number.Int64(); value != 42 {
		t.Fatalf("C2 = %d", value)
	}
	date, ok := row.Cell("D2")
	if !ok || date.Type != CellTypeDate {
		t.Fatalf("D2 = %+v", date)
	}
	if value, _ := date.Time(); !value.Equal(time.Date(2024, 1, 2, 12, 0, 0, 0, time.UTC)) {
		t.Fatalf("D2 = %v", value)
	}

	if !sheet.Next() {
		t.Fatalf("second Next() = false, err = %v", sheet.Err())
	}
	if sheet.Row().Index != 3 {
		t.Fatalf("row index = %d, want 3", sheet.Row().Index)
	}
	inline, ok := sheet.Row().Cell("A3")
	if !ok || inline.Type != CellTypeString {
		t.Fatalf("A3 = %+v", inline)
	}
	if value, _ := inline.String(); value != "inline" {
		t.Fatalf("A3 = %q", value)
	}
	boolean, ok := sheet.Row().Cell("B3")
	if !ok || boolean.Type != CellTypeBool {
		t.Fatalf("B3 = %+v", boolean)
	}
	if value, _ := boolean.Bool(); !value {
		t.Fatal("B3 = false, want true")
	}
}

func TestReader1904DateSystem(t *testing.T) {
	t.Parallel()

	parts := minimalPackage(`<row r="1"><c r="A1" s="1"><v>0</v></c></row>`)
	parts["xl/workbook.xml"] = xmlDeclaration + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<workbookPr date1904="1"/><sheets>` + sheetXML("Data", 1) + `</sheets></workbook>`
	path := writeRawXLSX(t, parts)
	reader, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer reader.Close()
	sheet, err := reader.OpenSheet("Data")
	if err != nil {
		t.Fatalf("OpenSheet() error = %v", err)
	}
	defer sheet.Close()
	if !sheet.Next() {
		t.Fatalf("Next() = false, err = %v", sheet.Err())
	}
	cell, ok := sheet.Row().Cell("A1")
	if !ok || cell.Type != CellTypeDate {
		t.Fatalf("A1 = %+v", cell)
	}
	if value, _ := cell.Time(); !value.Equal(time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)) {
		t.Fatalf("A1 = %v, want 1904-01-01", value)
	}
}

func TestReaderLimits(t *testing.T) {
	t.Parallel()

	path := writeRawXLSX(t, minimalPackage(`<row r="1"><c r="A1" t="inlineStr"><is><t>too long</t></is></c></row>`))

	reader, err := Open(path, WithReadMaxCellLength(3))
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer reader.Close()
	sheet, err := reader.OpenSheet("Data")
	if err != nil {
		t.Fatalf("OpenSheet() error = %v", err)
	}
	defer sheet.Close()
	if sheet.Next() {
		t.Fatal("Next() = true, want limit error")
	}
	if !errors.Is(sheet.Err(), ErrReadLimit) {
		t.Fatalf("Err() = %v, want ErrReadLimit", sheet.Err())
	}

	if _, err := Open(path, WithReadLimits(1<<20, 1)); !errors.Is(err, ErrReadLimit) {
		t.Fatalf("Open() error = %v, want ErrReadLimit", err)
	}
}

func TestReaderErrors(t *testing.T) {
	t.Parallel()

	dir := t.TempDir()
	plain := filepath.Join(dir, "plain.xlsx")
	if err := os.WriteFile(plain, []byte("not a zip"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := Open(plain); !errors.Is(err, ErrCorruptFile) {
		t.Fatalf("Open(plain) error = %v, want ErrCorruptFile", err)
	}

	ole := filepath.Join(dir, "ole.xlsx")
	if err := os.WriteFile(ole, append(append([]byte{}, oleSignature...), make([]byte, 64)...), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := Open(ole); !errors.Is(err, ErrUnsupportedFormat) || !errors.Is(err, ErrEncrypted) {
		t.Fatalf("Open(ole) error = %v, want ErrUnsupportedFormat and ErrEncrypted", err)
	}

	xlsb := writeRawXLSX(t, map[string]string{"xl/workbook.bin": "binary"})
	if _, err := Open(xlsb); !errors.Is(err, ErrUnsupportedFormat) {
		t.Fatalf("Open(xlsb) error = %v, want ErrUnsupportedFormat", err)
	}

	encrypted := writeRawXLSX(t, map[string]string{"EncryptedPackage": "encrypted"})
	if _, err := Open(encrypted); !errors.Is(err, ErrEncrypted) {
		t.Fatalf("Open(encrypted) error = %v, want ErrEncrypted", err)
	}

	zipSlip := writeRawXLSX(t, map[string]string{"../evil.xml": "x"})
	if _, err := Open(zipSlip); !errors.Is(err, ErrCorruptFile) {
		t.Fatalf("Open(zipSlip) error = %v, want ErrCorruptFile", err)
	}
}

func TestValidateBrokenPackages(t *testing.T) {
	t.Parallel()

	t.Run("valid", func(t *testing.T) {
		t.Parallel()
		path := writeRawXLSX(t, minimalPackage(`<row r="1"><c r="A1"><v>1</v></c></row>`))
		if err := Validate(path); err != nil {
			t.Fatalf("Validate() error = %v", err)
		}
	})

	t.Run("missing content types", func(t *testing.T) {
		t.Parallel()
		parts := minimalPackage(`<row r="1"><c r="A1"><v>1</v></c></row>`)
		delete(parts, "[Content_Types].xml")
		path := writeRawXLSX(t, parts)
		if err := Validate(path); err == nil {
			t.Fatal("Validate() = nil, want error")
		}
	})

	t.Run("missing worksheet part", func(t *testing.T) {
		t.Parallel()
		parts := minimalPackage(`<row r="1"><c r="A1"><v>1</v></c></row>`)
		delete(parts, "xl/worksheets/sheet1.xml")
		path := writeRawXLSX(t, parts)
		if err := Validate(path); !errors.Is(err, ErrCorruptFile) {
			t.Fatalf("Validate() error = %v, want ErrCorruptFile", err)
		}
	})

	t.Run("broken relationship", func(t *testing.T) {
		t.Parallel()
		parts := minimalPackage(`<row r="1"><c r="A1"><v>1</v></c></row>`)
		parts["xl/_rels/workbook.xml.rels"] = workbookRelsStart +
			relationshipXML(1, contentTypeWorksheetR, "worksheets/missing.xml") +
			relationshipXML(2, contentTypeStylesR, "styles.xml") +
			workbookRelsEnd
		path := writeRawXLSX(t, parts)
		if err := Validate(path); !errors.Is(err, ErrCorruptFile) {
			t.Fatalf("Validate() error = %v, want ErrCorruptFile", err)
		}
	})

	t.Run("zip slip", func(t *testing.T) {
		t.Parallel()
		path := writeRawXLSX(t, map[string]string{"../evil.xml": "x"})
		if err := Validate(path); !errors.Is(err, ErrCorruptFile) {
			t.Fatalf("Validate() error = %v, want ErrCorruptFile", err)
		}
	})

	t.Run("bad content type target", func(t *testing.T) {
		t.Parallel()
		parts := minimalPackage(`<row r="1"><c r="A1"><v>1</v></c></row>`)
		parts["[Content_Types].xml"] = contentTypesStart +
			contentTypeOverride("/xl/workbook.xml", contentTypeWorkbook) +
			contentTypeOverride("/xl/worksheets/sheet1.xml", contentTypeWorksheet) +
			contentTypeOverride("/xl/styles.xml", contentTypeStyles) +
			contentTypeOverride("/xl/missing.xml", contentTypeStyles) +
			contentTypesEnd
		path := writeRawXLSX(t, parts)
		if err := Validate(path); err == nil {
			t.Fatal("Validate() = nil, want error")
		}
	})

	t.Run("bad cell reference", func(t *testing.T) {
		t.Parallel()
		path := writeRawXLSX(t, minimalPackage(`<row r="1"><c r="BAD"><v>1</v></c></row>`))
		if err := Validate(path); !errors.Is(err, ErrCorruptFile) {
			t.Fatalf("Validate() error = %v, want ErrCorruptFile", err)
		}
	})

	t.Run("missing style", func(t *testing.T) {
		t.Parallel()
		path := writeRawXLSX(t, minimalPackage(`<row r="1"><c r="A1" s="99"><v>1</v></c></row>`))
		if err := Validate(path); !errors.Is(err, ErrCorruptFile) {
			t.Fatalf("Validate() error = %v, want ErrCorruptFile", err)
		}
	})

	t.Run("bad shared string index", func(t *testing.T) {
		t.Parallel()
		parts := minimalPackage(`<row r="1"><c r="A1" t="s"><v>4</v></c></row>`)
		parts["xl/sharedStrings.xml"] = xmlDeclaration + `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>only</t></si></sst>`
		parts["[Content_Types].xml"] = contentTypesStart +
			contentTypeOverride("/xl/workbook.xml", contentTypeWorkbook) +
			contentTypeOverride("/xl/worksheets/sheet1.xml", contentTypeWorksheet) +
			contentTypeOverride("/xl/styles.xml", contentTypeStyles) +
			contentTypeOverride("/xl/sharedStrings.xml", contentTypeSharedStrings) +
			contentTypesEnd
		parts["xl/_rels/workbook.xml.rels"] = workbookRelsStart +
			relationshipXML(1, contentTypeWorksheetR, "worksheets/sheet1.xml") +
			relationshipXML(2, contentTypeStylesR, "styles.xml") +
			relationshipXML(3, relationshipSharedStrings, "sharedStrings.xml") +
			workbookRelsEnd
		path := writeRawXLSX(t, parts)
		if err := Validate(path); !errors.Is(err, ErrCorruptFile) {
			t.Fatalf("Validate() error = %v, want ErrCorruptFile", err)
		}
	})
}

func TestValidateReaderAt(t *testing.T) {
	t.Parallel()

	path := writeRawXLSX(t, minimalPackage(`<row r="1"><c r="A1"><v>1</v></c></row>`))
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ValidateReaderAt(bytes.NewReader(data), int64(len(data))); err != nil {
		t.Fatalf("ValidateReaderAt() error = %v", err)
	}
}

func FuzzReader(f *testing.F) {
	f.Add([]byte("PK\x03\x04"))
	f.Add([]byte("not a zip"))
	f.Fuzz(func(_ *testing.T, data []byte) {
		if len(data) > 1<<20 {
			return
		}
		reader, err := OpenReaderAt(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			return
		}
		defer reader.Close()
		for _, sheet := range reader.Sheets() {
			sheetReader, err := reader.OpenSheet(sheet.Name)
			if err != nil {
				continue
			}
			for sheetReader.Next() {
				_ = sheetReader.Row()
			}
			_ = sheetReader.Err()
			_ = sheetReader.Close()
		}
	})
}

func TestReaderMultipleBatches(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "batches.xlsx")
	workbook := NewWorkbook(path)
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	for i := 0; i < 25; i++ {
		row, err := sheet.AddRow()
		if err != nil {
			t.Fatal(err)
		}
		if _, err := row.WriteInt(int64(i)); err != nil {
			t.Fatal(err)
		}
		if err := sheet.CommitRows(); err != nil {
			t.Fatal(err)
		}
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}

	reader, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reader.Close()
	stream, err := reader.OpenSheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	defer stream.Close()
	count := 0
	for stream.Next() {
		row := stream.Row()
		cell, ok := row.CellAt(0)
		if !ok {
			t.Fatalf("row %d has no cell", row.Index)
		}
		value, _ := cell.Int64()
		if value != int64(count) {
			t.Fatalf("row %d = %d, want %d", row.Index, value, count)
		}
		count++
	}
	if err := stream.Err(); err != nil {
		t.Fatal(err)
	}
	if count != 25 {
		t.Fatalf("read %d rows, want 25", count)
	}
}

func TestLargeStreamingWorkbook(t *testing.T) {
	t.Parallel()

	const rows = 5_000
	path := filepath.Join(t.TempDir(), "large.xlsx")
	workbook := NewWorkbook(path)
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	for i := 0; i < rows; i++ {
		row, err := sheet.AddRow()
		if err != nil {
			t.Fatal(err)
		}
		if _, err := row.WriteInt(int64(i)); err != nil {
			t.Fatal(err)
		}
		if _, err := row.WriteString("some repeated value"); err != nil {
			t.Fatal(err)
		}
		if (i+1)%500 == 0 {
			if err := sheet.CommitRows(); err != nil {
				t.Fatal(err)
			}
		}
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}

	reader, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reader.Close()
	stream, err := reader.OpenSheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	defer stream.Close()
	count := 0
	for stream.Next() {
		count++
	}
	if err := stream.Err(); err != nil {
		t.Fatal(err)
	}
	if count != rows {
		t.Fatalf("read %d rows, want %d", count, rows)
	}
}
