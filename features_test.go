package xlsx

import (
	"archive/zip"
	"bytes"
	"context"
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"
)

func TestWriterOptionsAndTempDir(t *testing.T) {
	t.Parallel()

	base := t.TempDir()
	tempDir := filepath.Join(base, "nested-temp")
	if err := os.Mkdir(tempDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "workbook.xlsx")

	workbook, err := NewWorkbookWithOptions(path, Options{
		TempDir:       tempDir,
		BufferSize:    128,
		TempDirMode:   0o700,
		FileMode:      0o600,
		Compression:   CompressionStore,
		KeepTempFiles: true,
	})
	if err != nil {
		t.Fatalf("NewWorkbookWithOptions() error = %v", err)
	}
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	row, err := sheet.AddRow()
	if err != nil {
		t.Fatal(err)
	}
	if _, err := row.WriteString("value"); err != nil {
		t.Fatal(err)
	}
	if err := sheet.CommitRows(); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}

	info, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	if runtime.GOOS != "windows" {
		if got := info.Mode().Perm(); got != 0o600 {
			t.Fatalf("output mode = %o, want 600", got)
		}
	}

	tempRoot := workbook.TempDir()
	if tempRoot == "" {
		t.Fatal("TempDir() is empty")
	}
	if !strings.HasPrefix(tempRoot, tempDir) {
		t.Fatalf("TempDir() = %q, want under %q", tempRoot, tempDir)
	}
	if _, err := os.Stat(tempRoot); err != nil {
		t.Fatalf("kept temp dir does not exist: %v", err)
	}
	if err := workbook.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}
	if err := workbook.Close(); err != nil {
		t.Fatalf("second Close() error = %v", err)
	}
	_ = os.RemoveAll(tempRoot)

	if _, err := NewWorkbookWithOptions(path, Options{TempDir: filepath.Join(base, "missing")}); err == nil {
		t.Fatal("NewWorkbookWithOptions() accepted a missing temp dir")
	}
}

func TestNewWriter(t *testing.T) {
	t.Parallel()

	var buf bytes.Buffer
	workbook, err := NewWriter(&buf, Options{SharedStrings: true})
	if err != nil {
		t.Fatalf("NewWriter() error = %v", err)
	}
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	row, err := sheet.AddRow()
	if err != nil {
		t.Fatal(err)
	}
	if err := row.Append("hello", 42, true); err != nil {
		t.Fatal(err)
	}
	if err := sheet.CommitRows(); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	if buf.Len() == 0 {
		t.Fatal("output buffer is empty")
	}

	reader, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader() error = %v", err)
	}
	names := make(map[string]bool)
	for _, file := range reader.File {
		names[file.Name] = true
	}
	for _, name := range []string{"[Content_Types].xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/sharedStrings.xml"} {
		if !names[name] {
			t.Fatalf("missing part %q", name)
		}
	}
}

func TestWriterStylesAndSharedStrings(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "styled.xlsx")
	workbook, err := NewWorkbookWithOptions(path, Options{SharedStrings: true})
	if err != nil {
		t.Fatal(err)
	}
	styleID, err := workbook.AddStyle(Style{
		Bold:         true,
		Italic:       true,
		Underline:    true,
		FontSize:     14,
		FontColor:    "#FF0000",
		FillColor:    "00FF00",
		NumberFormat: "0.000",
		Horizontal:   "right",
		Vertical:     "center",
		WrapText:     true,
	})
	if err != nil {
		t.Fatalf("AddStyle() error = %v", err)
	}
	if styleID != 2 {
		t.Fatalf("first style ID = %d, want 2", styleID)
	}

	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	for i := 0; i < 2; i++ {
		row, err := sheet.AddRow()
		if err != nil {
			t.Fatal(err)
		}
		cell, _ := row.WriteString("repeated")
		cell.SetStyle(styleID)
		if _, err := row.WriteInt(42); err != nil {
			t.Fatal(err)
		}
		if _, err := row.WriteTime(time.Date(2024, 1, 2, 3, 4, 5, 0, time.UTC)); err != nil {
			t.Fatal(err)
		}
	}
	if err := sheet.CommitRows(); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}

	parts := readZipParts(t, path)
	shared := string(parts["xl/sharedStrings.xml"].Data)
	if !strings.Contains(shared, `uniqueCount="1"`) || !strings.Contains(shared, `count="2"`) {
		t.Fatalf("unexpected shared strings XML: %s", shared)
	}
	if !strings.Contains(shared, "repeated") {
		t.Fatalf("shared strings XML does not contain the value: %s", shared)
	}

	styles := string(parts["xl/styles.xml"].Data)
	for _, fragment := range []string{`<b/>`, `<i/>`, `<u/>`, `<color rgb="FFFF0000"/>`, `<fill>`, `formatCode="0.000"`, `horizontal="right"`, `wrapText="1"`} {
		if !strings.Contains(styles, fragment) {
			t.Errorf("styles.xml does not contain %q", fragment)
		}
	}

	sheetXML := string(parts["xl/worksheets/sheet1.xml"].Data)
	if !strings.Contains(sheetXML, `t="s"`) {
		t.Error("worksheet does not use shared strings")
	}
	if !strings.Contains(sheetXML, `s="2"`) {
		t.Error("worksheet does not reference the user style")
	}
	if !strings.Contains(sheetXML, `s="1"`) {
		t.Error("worksheet does not reference the date style")
	}
}

func TestWriterHyperlinksMergesFreezeAndAutoFilter(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "layout.xlsx")
	workbook := NewWorkbook(path)
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := sheet.FreezePanes(1, 1); err != nil {
		t.Fatalf("FreezePanes() error = %v", err)
	}
	if err := sheet.SetAutoFilter("A1:C3"); err != nil {
		t.Fatalf("SetAutoFilter() error = %v", err)
	}
	if err := sheet.MergeCells("A4:B4"); err != nil {
		t.Fatalf("MergeCells() error = %v", err)
	}

	row, err := sheet.AddRow()
	if err != nil {
		t.Fatal(err)
	}
	cell, _ := row.WriteString("external")
	cell.SetHyperlink("https://example.com?a=1&b=2")
	internal, _ := row.WriteString("internal")
	internal.SetHyperlink("#Summary!A1")
	if _, err := row.WriteString("plain"); err != nil {
		t.Fatal(err)
	}
	if err := sheet.CommitRows(); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}

	parts := readZipParts(t, path)
	sheetXML := string(parts["xl/worksheets/sheet1.xml"].Data)
	for _, fragment := range []string{
		`<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>`,
		`<autoFilter ref="A1:C3"/>`,
		`<mergeCells count="1"><mergeCell ref="A4:B4"/></mergeCells>`,
		`<hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="B1" r:id="rId2"/></hyperlinks>`,
	} {
		if !strings.Contains(sheetXML, fragment) {
			t.Errorf("worksheet XML does not contain %q\n%s", fragment, sheetXML)
		}
	}

	rels := string(parts["xl/worksheets/_rels/sheet1.xml.rels"].Data)
	if !strings.Contains(rels, `Target="https://example.com?a=1&amp;b=2" TargetMode="External"`) {
		t.Errorf("worksheet rels do not contain the external hyperlink: %s", rels)
	}
	if !strings.Contains(rels, `Target="#Summary!A1"`) || strings.Contains(rels, `Target="#Summary!A1" TargetMode`) {
		t.Errorf("internal hyperlink should not be external: %s", rels)
	}
}

func TestWriterFormulaAndRowHeight(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "formula.xlsx")
	workbook := NewWorkbook(path)
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	row, err := sheet.AddRow()
	if err != nil {
		t.Fatal(err)
	}
	if err := row.SetHeight(21); err != nil {
		t.Fatal(err)
	}
	if _, err := row.WriteInt(2); err != nil {
		t.Fatal(err)
	}
	formula, _ := row.AddCell()
	formula.SetFormula("=A1*2")
	if err := sheet.CommitRows(); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}

	parts := readZipParts(t, path)
	sheetXML := string(parts["xl/worksheets/sheet1.xml"].Data)
	if !strings.Contains(sheetXML, `<row r="1" ht="21" customHeight="1">`) {
		t.Errorf("row height is missing: %s", sheetXML)
	}
	if !strings.Contains(sheetXML, `<f>A1*2</f>`) {
		t.Errorf("formula is missing or not normalized: %s", sheetXML)
	}
	if !strings.Contains(string(parts["xl/workbook.xml"].Data), `fullCalcOnLoad="1"`) {
		t.Error("workbook does not request formula recalculation")
	}
}

func TestWriterContextCancellation(t *testing.T) {
	t.Parallel()

	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	workbook, err := NewWorkbookWithOptions(filepath.Join(t.TempDir(), "canceled.xlsx"), Options{Context: ctx})
	if err != nil {
		t.Fatal(err)
	}
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if _, err := sheet.AddRow(); !errors.Is(err, context.Canceled) {
		t.Fatalf("AddRow() error = %v, want context.Canceled", err)
	}
	if err := workbook.Close(); err != nil {
		t.Fatal(err)
	}

	ctx, cancel = context.WithCancel(context.Background())
	workbook, err = NewWorkbookWithOptions(filepath.Join(t.TempDir(), "canceled-mid.xlsx"), Options{Context: ctx})
	if err != nil {
		t.Fatal(err)
	}
	sheet = workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	row, err := sheet.AddRow()
	if err != nil {
		t.Fatal(err)
	}
	if _, err := row.WriteString("value"); err != nil {
		t.Fatal(err)
	}
	cancel()
	if err := sheet.CommitRows(); !errors.Is(err, context.Canceled) {
		t.Fatalf("CommitRows() error = %v, want context.Canceled", err)
	}
	if err := workbook.Close(); err != nil {
		t.Fatal(err)
	}
}

func TestWriterLimitsAndValidation(t *testing.T) {
	t.Parallel()

	t.Run("cell length", func(t *testing.T) {
		t.Parallel()
		workbook, err := NewWorkbookWithOptions(filepath.Join(t.TempDir(), "long.xlsx"), Options{MaxCellLength: 3})
		if err != nil {
			t.Fatal(err)
		}
		sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
		row, _ := sheet.AddRow()
		if _, err := row.WriteString("long"); err != nil {
			t.Fatal(err)
		}
		if err := sheet.CommitRows(); !errors.Is(err, ErrCellTooLong) {
			t.Fatalf("CommitRows() error = %v, want ErrCellTooLong", err)
		}
		_ = workbook.Close()
	})

	t.Run("sheet limit", func(t *testing.T) {
		t.Parallel()
		workbook, err := NewWorkbookWithOptions(filepath.Join(t.TempDir(), "sheets.xlsx"), Options{MaxSheets: 1})
		if err != nil {
			t.Fatal(err)
		}
		workbook.AddWorksheet(&WorksheetOptions{Name: "One"})
		second := workbook.AddWorksheet(&WorksheetOptions{Name: "Two"})
		if err := second.Commit(); !errors.Is(err, ErrTooManySheets) {
			t.Fatalf("Commit() error = %v, want ErrTooManySheets", err)
		}
		_ = workbook.Close()
	})

	t.Run("styles", func(t *testing.T) {
		t.Parallel()
		workbook := NewWorkbook(filepath.Join(t.TempDir(), "styles.xlsx"))
		for _, style := range []Style{
			{FontSize: 500},
			{FontColor: "nope"},
			{Horizontal: "sideways"},
			{NumberFormat: "bad\x00format"},
		} {
			if _, err := workbook.AddStyle(style); !errors.Is(err, ErrInvalidStyle) {
				t.Fatalf("AddStyle(%+v) error = %v, want ErrInvalidStyle", style, err)
			}
		}
		sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
		row, _ := sheet.AddRow()
		cell, _ := row.WriteString("x")
		cell.SetStyle(99)
		if err := sheet.CommitRows(); !errors.Is(err, ErrInvalidStyle) {
			t.Fatalf("CommitRows() error = %v, want ErrInvalidStyle", err)
		}
		_ = workbook.Close()
	})

	t.Run("formula and hyperlink", func(t *testing.T) {
		t.Parallel()
		workbook := NewWorkbook(filepath.Join(t.TempDir(), "invalid.xlsx"))
		sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
		row, _ := sheet.AddRow()
		cell, _ := row.AddCell()
		cell.Formula = " "
		if err := sheet.CommitRows(); !errors.Is(err, ErrInvalidFormula) {
			t.Fatalf("CommitRows() error = %v, want ErrInvalidFormula", err)
		}
		_ = workbook.Close()
	})

	t.Run("merge ranges", func(t *testing.T) {
		t.Parallel()
		workbook := NewWorkbook(filepath.Join(t.TempDir(), "merge.xlsx"))
		sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
		if err := sheet.MergeCells("B2:A1"); !errors.Is(err, ErrInvalidRange) {
			t.Fatalf("MergeCells() error = %v, want ErrInvalidRange", err)
		}
		if err := sheet.MergeCells("A1:B2"); err != nil {
			t.Fatal(err)
		}
		if err := sheet.MergeCells("B2:C3"); !errors.Is(err, ErrInvalidRange) {
			t.Fatalf("overlapping MergeCells() error = %v, want ErrInvalidRange", err)
		}
		_ = workbook.Close()
	})

	t.Run("row height", func(t *testing.T) {
		t.Parallel()
		workbook := NewWorkbook(filepath.Join(t.TempDir(), "height.xlsx"))
		sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
		row, _ := sheet.AddRow()
		if err := row.SetHeight(500); !errors.Is(err, ErrInvalidRow) {
			t.Fatalf("SetHeight() error = %v, want ErrInvalidRow", err)
		}
		row.Height = 500
		if _, err := row.WriteString("x"); err != nil {
			t.Fatal(err)
		}
		if err := sheet.CommitRows(); !errors.Is(err, ErrInvalidRow) {
			t.Fatalf("CommitRows() error = %v, want ErrInvalidRow", err)
		}
		_ = workbook.Close()
	})
}

func TestGoldenPackageStructure(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "golden.xlsx")
	workbook := NewWorkbook(path)
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	row, err := sheet.AddRow()
	if err != nil {
		t.Fatal(err)
	}
	if _, err := row.WriteString("value"); err != nil {
		t.Fatal(err)
	}
	if err := sheet.CommitRows(); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}

	parts := readZipParts(t, path)
	want := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"xl/_rels/workbook.xml.rels",
		"xl/styles.xml",
		"xl/workbook.xml",
		"xl/worksheets/sheet1.xml",
	}
	if len(parts) != len(want) {
		t.Fatalf("package has %d parts, want %d: %v", len(parts), len(want), parts)
	}
	for _, name := range want {
		part, ok := parts[name]
		if !ok {
			t.Fatalf("missing golden part %q", name)
		}
		assertWellFormedXML(t, name, part.Data)
		if !strings.HasPrefix(string(part.Data), xmlDeclaration) {
			t.Errorf("%s does not start with the XML declaration", name)
		}
	}

	// These are the stable structural anchors that consumers rely on.
	workbookXML := string(parts["xl/workbook.xml"].Data)
	if !strings.Contains(workbookXML, `<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>`) {
		t.Errorf("unexpected workbook sheet markup: %s", workbookXML)
	}
	rels := string(parts["xl/_rels/workbook.xml.rels"].Data)
	if !strings.Contains(rels, `Target="worksheets/sheet1.xml"`) || !strings.Contains(rels, `Target="styles.xml"`) {
		t.Errorf("unexpected workbook relationships: %s", rels)
	}
}

func TestOptionsValidation(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "options.xlsx")
	for name, opts := range map[string]Options{
		"negative buffer":      {BufferSize: -1},
		"negative cell length": {MaxCellLength: -1},
		"too many rows":        {MaxRows: maxWorksheetRows + 1},
		"too many columns":     {MaxColumns: maxWorksheetColumns + 1},
		"negative flush":       {FlushEveryRows: -1},
		"unknown compression":  {Compression: 99},
	} {
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			if _, err := NewWorkbookWithOptions(path, opts); err == nil {
				t.Fatalf("NewWorkbookWithOptions(%+v) = nil error", opts)
			}
		})
	}
	if _, err := NewWriter(nil, Options{}); err == nil {
		t.Fatal("NewWriter(nil) = nil error")
	}
}

func TestReadOptionsValidation(t *testing.T) {
	t.Parallel()

	path := writeRawXLSX(t, minimalPackage(`<row r="1"><c r="A1"><v>1</v></c></row>`))
	if _, err := Open(path, WithReadMaxCellLength(-1)); err == nil {
		t.Fatal("Open() with a negative cell limit returned nil error")
	}
	if _, err := Open(path, WithReadMaxRows(-1)); err == nil {
		t.Fatal("Open() with a negative row limit returned nil error")
	}
}
