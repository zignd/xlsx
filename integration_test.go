package xlsx

import (
	"archive/zip"
	"bytes"
	"database/sql"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"testing"
	"time"
)

type zipPart struct {
	Name string
	Data []byte
}

type parsedCell struct {
	R  string `xml:"r,attr"`
	T  string `xml:"t,attr"`
	S  string `xml:"s,attr"`
	V  string `xml:"v"`
	IS struct {
		T string `xml:"t"`
	} `xml:"is"`
}

type parsedRow struct {
	R     int          `xml:"r,attr"`
	Cells []parsedCell `xml:"c"`
}

type parsedWorksheet struct {
	Columns struct {
		Items []struct {
			Min   int     `xml:"min,attr"`
			Max   int     `xml:"max,attr"`
			Width float64 `xml:"width,attr"`
		} `xml:"col"`
	} `xml:"cols"`
	Rows []parsedRow `xml:"sheetData>row"`
}

func readZipParts(t *testing.T, path string) map[string]zipPart {
	t.Helper()

	reader, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("zip.OpenReader(%q) error = %v", path, err)
	}
	defer reader.Close()

	parts := make(map[string]zipPart, len(reader.File))
	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			t.Fatalf("open zip entry %q: %v", file.Name, err)
		}
		data, err := io.ReadAll(rc)
		closeErr := rc.Close()
		if err != nil {
			t.Fatalf("read zip entry %q: %v", file.Name, err)
		}
		if closeErr != nil {
			t.Fatalf("close zip entry %q: %v", file.Name, closeErr)
		}
		parts[file.Name] = zipPart{Name: file.Name, Data: data}
	}
	return parts
}

func assertWellFormedXML(t *testing.T, name string, data []byte) {
	t.Helper()

	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		_, err := decoder.Token()
		if errors.Is(err, io.EOF) {
			return
		}
		if err != nil {
			t.Fatalf("%s is not well-formed XML: %v", name, err)
		}
	}
}

func parseWorksheet(t *testing.T, name string, data []byte) parsedWorksheet {
	t.Helper()

	var sheet parsedWorksheet
	if err := xml.Unmarshal(data, &sheet); err != nil {
		t.Fatalf("failed to parse %s: %v", name, err)
	}
	return sheet
}

func cellByReference(sheet parsedWorksheet, reference string) (parsedCell, bool) {
	for _, row := range sheet.Rows {
		for _, cell := range row.Cells {
			if cell.R == reference {
				return cell, true
			}
		}
	}
	return parsedCell{}, false
}

func TestWorkbookEndToEnd(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "report.xlsx")
	wb := NewWorkbook(path)

	dataSheet := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := dataSheet.DefineColumns([]*WorksheetColumn{
		{Key: "name", Value: "Name", Width: 24},
		{Key: "amount", Value: "Amount"},
		{Key: "when", Value: "When"},
		{Key: "active", Value: "Active"},
	}); err != nil {
		t.Fatalf("DefineColumns() error = %v", err)
	}

	when := time.Date(2024, 1, 2, 3, 4, 5, 0, time.UTC)
	for i := 0; i < 3; i++ {
		row, err := dataSheet.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}

		name, err := row.AddCellWithKey("name")
		if err != nil {
			t.Fatalf("AddCellWithKey(name) error = %v", err)
		}
		name.Value = fmt.Sprintf("item & <%d>", i)

		if i != 2 {
			amount, err := row.AddCellWithKey("amount")
			if err != nil {
				t.Fatalf("AddCellWithKey(amount) error = %v", err)
			}
			amount.Value = 1.5 * float64(i+1)
		}

		dateCell, err := row.AddCellWithKey("when")
		if err != nil {
			t.Fatalf("AddCellWithKey(when) error = %v", err)
		}
		dateCell.Value = when

		active, err := row.AddCellWithKey("active")
		if err != nil {
			t.Fatalf("AddCellWithKey(active) error = %v", err)
		}
		active.Value = i%2 == 0
	}
	if err := dataSheet.CommitRows(); err != nil {
		t.Fatalf("CommitRows() error = %v", err)
	}
	if err := dataSheet.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}

	summary := wb.AddWorksheet(&WorksheetOptions{Name: "Summary"})
	row, err := summary.AddRow()
	if err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	values := []interface{}{
		"total",
		42,
		int8(-7),
		3.25,
		float32(0.1),
		true,
		when,
		nil,
		sql.NullString{String: "nullable", Valid: true},
		[]byte("bytes"),
	}
	for _, value := range values {
		cell, err := row.AddCell()
		if err != nil {
			t.Fatalf("AddCell() error = %v", err)
		}
		cell.Value = value
	}
	if err := summary.CommitRows(); err != nil {
		t.Fatalf("summary CommitRows() error = %v", err)
	}

	tempRoot := wb.tempRootDir
	if err := wb.Commit(); err != nil {
		t.Fatalf("workbook Commit() error = %v", err)
	}
	if _, err := os.Stat(tempRoot); !os.IsNotExist(err) {
		t.Fatalf("temporary directory %q still exists after commit (stat error = %v)", tempRoot, err)
	}

	parts := readZipParts(t, path)
	for _, name := range []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"xl/workbook.xml",
		"xl/_rels/workbook.xml.rels",
		"xl/styles.xml",
		"xl/worksheets/sheet1.xml",
		"xl/worksheets/sheet2.xml",
	} {
		part, ok := parts[name]
		if !ok {
			t.Fatalf("missing zip part %q", name)
		}
		assertWellFormedXML(t, name, part.Data)
	}

	var contentTypes struct {
		Overrides []struct {
			PartName    string `xml:"PartName,attr"`
			ContentType string `xml:"ContentType,attr"`
		} `xml:"Override"`
	}
	if err := xml.Unmarshal(parts["[Content_Types].xml"].Data, &contentTypes); err != nil {
		t.Fatalf("parse content types: %v", err)
	}
	overrides := make(map[string]string, len(contentTypes.Overrides))
	for _, override := range contentTypes.Overrides {
		overrides[override.PartName] = override.ContentType
	}
	for partName, want := range map[string]string{
		"/xl/workbook.xml":          contentTypeWorkbook,
		"/xl/worksheets/sheet1.xml": contentTypeWorksheet,
		"/xl/worksheets/sheet2.xml": contentTypeWorksheet,
		"/xl/styles.xml":            contentTypeStyles,
	} {
		if got := overrides[partName]; got != want {
			t.Errorf("content type for %q = %q, want %q", partName, got, want)
		}
	}

	var workbook struct {
		Sheets []struct {
			Name string `xml:"name,attr"`
			ID   int    `xml:"sheetId,attr"`
		} `xml:"sheets>sheet"`
	}
	if err := xml.Unmarshal(parts["xl/workbook.xml"].Data, &workbook); err != nil {
		t.Fatalf("parse workbook: %v", err)
	}
	if got, want := len(workbook.Sheets), 2; got != want {
		t.Fatalf("workbook has %d sheets, want %d", got, want)
	}
	if workbook.Sheets[0].Name != "Data" || workbook.Sheets[1].Name != "Summary" {
		t.Fatalf("unexpected sheet names: %+v", workbook.Sheets)
	}

	var relationships struct {
		Items []struct {
			ID     string `xml:"Id,attr"`
			Type   string `xml:"Type,attr"`
			Target string `xml:"Target,attr"`
		} `xml:"Relationship"`
	}
	if err := xml.Unmarshal(parts["xl/_rels/workbook.xml.rels"].Data, &relationships); err != nil {
		t.Fatalf("parse workbook relationships: %v", err)
	}
	if got, want := len(relationships.Items), 3; got != want {
		t.Fatalf("workbook has %d relationships, want %d", got, want)
	}
	if got := relationships.Items[2].Target; got != "styles.xml" {
		t.Fatalf("styles relationship target = %q, want styles.xml", got)
	}

	sheet := parseWorksheet(t, "sheet1.xml", parts["xl/worksheets/sheet1.xml"].Data)
	if got, want := len(sheet.Rows), 4; got != want {
		t.Fatalf("sheet1 has %d rows, want %d", got, want)
	}
	for i, row := range sheet.Rows {
		if got, want := row.R, i+1; got != want {
			t.Fatalf("row %d has r=%d, want %d", i, got, want)
		}
	}

	if got, want := len(sheet.Columns.Items), 1; got != want {
		t.Fatalf("sheet1 has %d column definitions, want %d", got, want)
	}
	if got, want := sheet.Columns.Items[0].Width, 24.0; got != want {
		t.Fatalf("column width = %v, want %v", got, want)
	}

	headerName, ok := cellByReference(sheet, "A1")
	if !ok || headerName.IS.T != "Name" {
		t.Fatalf("A1 = %+v, want header Name", headerName)
	}
	nameCell, ok := cellByReference(sheet, "A2")
	if !ok || nameCell.IS.T != "item & <0>" {
		t.Fatalf("A2 = %+v, want escaped text item & <0>", nameCell)
	}
	amountCell, ok := cellByReference(sheet, "B2")
	if !ok || amountCell.V != "1.5" {
		t.Fatalf("B2 = %+v, want number 1.5", amountCell)
	}
	dateValue, ok := cellByReference(sheet, "C2")
	if !ok {
		t.Fatal("C2 is missing")
	}
	gotDate, err := strconv.ParseFloat(dateValue.V, 64)
	if err != nil {
		t.Fatalf("C2 value %q is not a number: %v", dateValue.V, err)
	}
	if want := timeToExcelTime(when); gotDate != want {
		t.Fatalf("C2 serial = %v, want %v", gotDate, want)
	}
	if dateValue.S != "1" {
		t.Fatalf("C2 style = %q, want 1", dateValue.S)
	}
	activeCell, ok := cellByReference(sheet, "D2")
	if !ok || activeCell.T != "b" || activeCell.V != "1" {
		t.Fatalf("D2 = %+v, want boolean true", activeCell)
	}

	blankAmount, ok := cellByReference(sheet, "B4")
	if !ok {
		t.Fatal("B4 is missing")
	}
	if blankAmount.V != "" || blankAmount.IS.T != "" {
		t.Fatalf("B4 = %+v, want blank cell", blankAmount)
	}

	summarySheet := parseWorksheet(t, "sheet2.xml", parts["xl/worksheets/sheet2.xml"].Data)
	if got, want := len(summarySheet.Rows), 1; got != want {
		t.Fatalf("summary has %d rows, want %d", got, want)
	}
	expectedText := map[string]string{
		"A1": "total",
		"H1": "",
		"I1": "nullable",
		"J1": "bytes",
	}
	for reference, want := range expectedText {
		cell, ok := cellByReference(summarySheet, reference)
		if !ok {
			t.Fatalf("%s is missing", reference)
		}
		if want == "" {
			if cell.IS.T != "" || cell.V != "" {
				t.Fatalf("%s = %+v, want a blank cell", reference, cell)
			}
			continue
		}
		if got := cell.IS.T; got != want {
			t.Fatalf("%s text = %q, want %q", reference, got, want)
		}
	}
	expectedNumbers := map[string]string{
		"B1": "42",
		"C1": "-7",
		"D1": "3.25",
		"E1": "0.1",
	}
	for reference, want := range expectedNumbers {
		cell, ok := cellByReference(summarySheet, reference)
		if !ok {
			t.Fatalf("%s is missing", reference)
		}
		if got := cell.V; got != want {
			t.Fatalf("%s = %q, want %q", reference, got, want)
		}
	}
}

func TestMultipleCommitRowsBatches(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "batches.xlsx")
	wb := NewWorkbook(path)
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Batches"})

	for i := 0; i < 5; i++ {
		row, err := ws.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		cell, err := row.AddCell()
		if err != nil {
			t.Fatalf("AddCell() error = %v", err)
		}
		cell.Value = i
		if err := ws.CommitRows(); err != nil {
			t.Fatalf("CommitRows() error = %v", err)
		}
	}

	if err := wb.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}

	parts := readZipParts(t, path)
	sheet := parseWorksheet(t, "sheet1.xml", parts["xl/worksheets/sheet1.xml"].Data)
	if got, want := len(sheet.Rows), 5; got != want {
		t.Fatalf("sheet has %d rows, want %d", got, want)
	}
	for i, row := range sheet.Rows {
		if got, want := row.R, i+1; got != want {
			t.Fatalf("row %d has r=%d, want %d", i, got, want)
		}
		if got, want := row.Cells[0].V, strconv.Itoa(i); got != want {
			t.Fatalf("row %d cell = %q, want %q", i, got, want)
		}
	}
}

func TestCommitOverwritesExistingFile(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "workbook.xlsx")
	if err := os.WriteFile(path, []byte("not a zip file"), 0o644); err != nil {
		t.Fatalf("WriteFile() error = %v", err)
	}

	wb := NewWorkbook(path)
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	row, err := ws.AddRow()
	if err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	cell, err := row.AddCell()
	if err != nil {
		t.Fatalf("AddCell() error = %v", err)
	}
	cell.Value = "hello"

	if err := ws.CommitRows(); err != nil {
		t.Fatalf("CommitRows() error = %v", err)
	}
	if err := wb.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}

	parts := readZipParts(t, path)
	if _, ok := parts["xl/worksheets/sheet1.xml"]; !ok {
		t.Fatal("overwritten file is not a valid workbook")
	}
}

func TestCommitCreatesOutputDirectory(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "nested", "deeper", "workbook.xlsx")
	wb := NewWorkbook(path)
	wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := wb.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	if _, err := os.Stat(path); err != nil {
		t.Fatalf("output file does not exist: %v", err)
	}
}
