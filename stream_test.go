package xlsx

import (
	"bytes"
	"errors"
	"io"
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"testing"
	"time"
)

func streamedWorkbookBytes(t *testing.T, opts Options, build func(*Workbook) error) []byte {
	t.Helper()

	var buf bytes.Buffer
	workbook, err := NewStreamWriter(&buf, opts)
	if err != nil {
		t.Fatalf("NewStreamWriter() error = %v", err)
	}
	if err := build(workbook); err != nil {
		t.Fatalf("build() error = %v", err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	return buf.Bytes()
}

func TestStreamWriterGeneratedWorkbook(t *testing.T) {
	t.Parallel()

	data := streamedWorkbookBytes(t, Options{SharedStrings: true}, func(workbook *Workbook) error {
		styleID, err := workbook.AddStyle(Style{Bold: true, NumberFormat: "#,##0.00"})
		if err != nil {
			return err
		}

		sales := workbook.AddWorksheet(&WorksheetOptions{Name: "Sales"})
		if err := sales.FreezePanes(1, 0); err != nil {
			return err
		}
		if err := sales.DefineColumns([]*WorksheetColumn{
			{Key: "name", Value: "Name", Width: 20},
			{Key: "amount", Value: "Amount"},
			{Key: "when", Value: "When"},
		}); err != nil {
			return err
		}
		if err := sales.SetAutoFilter("A1:C2"); err != nil {
			return err
		}
		row, err := sales.AddRow()
		if err != nil {
			return err
		}
		name, _ := row.AddCellWithKey("name")
		name.SetString("Widget")
		name.SetHyperlink("https://example.com")
		amount, _ := row.AddCellWithKey("amount")
		amount.SetFloat(12.5).SetStyle(styleID)
		when, _ := row.AddCellWithKey("when")
		when.SetTime(time.Date(2024, 1, 2, 3, 4, 5, 0, time.UTC))
		if err := sales.CommitRows(); err != nil {
			return err
		}
		if err := sales.MergeCells("A3:C3"); err != nil {
			return err
		}
		if err := sales.Commit(); err != nil {
			return err
		}

		summary := workbook.AddWorksheet(&WorksheetOptions{Name: "Summary"})
		summaryRow, err := summary.AddRow()
		if err != nil {
			return err
		}
		if _, err := summaryRow.WriteString("Total"); err != nil {
			return err
		}
		total, err := summaryRow.AddCell()
		if err != nil {
			return err
		}
		total.SetFormula("=SUM(Sales!B2:B2)")
		return summary.CommitRows()
	})

	if err := ValidateReaderAt(bytes.NewReader(data), int64(len(data))); err != nil {
		t.Fatalf("ValidateReaderAt() error = %v", err)
	}

	reader, err := OpenReaderAt(bytes.NewReader(data), int64(len(data)), WithSharedStringsMode(SharedStringsDisk))
	if err != nil {
		t.Fatalf("OpenReaderAt() error = %v", err)
	}
	defer reader.Close()

	sheets := reader.Sheets()
	if len(sheets) != 2 || sheets[0].Name != "Sales" || sheets[1].Name != "Summary" {
		t.Fatalf("unexpected sheets: %+v", sheets)
	}
	sheet, err := reader.OpenSheet("Sales")
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
	if value, _ := mustCell(t, sheet.Row(), "A2").String(); value != "Widget" {
		t.Fatalf("A2 = %q, want Widget", value)
	}
	if value, _ := mustCell(t, sheet.Row(), "B2").Float64(); value != 12.5 {
		t.Fatalf("B2 = %v, want 12.5", value)
	}
}

func mustCell(t *testing.T, row *RowData, ref string) *CellData {
	t.Helper()
	cell, ok := row.Cell(ref)
	if !ok {
		t.Fatalf("cell %s is missing", ref)
	}
	return cell
}

func TestStreamWriterIsSequential(t *testing.T) {
	t.Parallel()

	var buf bytes.Buffer
	workbook, err := NewStreamWriter(&buf, Options{})
	if err != nil {
		t.Fatal(err)
	}
	first := workbook.AddWorksheet(&WorksheetOptions{Name: "First"})
	second := workbook.AddWorksheet(&WorksheetOptions{Name: "Second"})

	row, _ := first.AddRow()
	if _, err := row.WriteString("first"); err != nil {
		t.Fatal(err)
	}
	if err := first.CommitRows(); err != nil {
		t.Fatal(err)
	}

	row, _ = second.AddRow()
	if _, err := row.WriteString("second"); err != nil {
		t.Fatal(err)
	}
	if err := second.CommitRows(); !errors.Is(err, ErrStreamingOrder) {
		t.Fatalf("second CommitRows() error = %v, want ErrStreamingOrder", err)
	}

	if err := first.Commit(); err != nil {
		t.Fatal(err)
	}
	if err := second.CommitRows(); err != nil {
		t.Fatalf("second CommitRows() after first Commit() error = %v", err)
	}
	if err := second.Commit(); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}
	if err := ValidateReaderAt(bytes.NewReader(buf.Bytes()), int64(buf.Len())); err != nil {
		t.Fatalf("ValidateReaderAt() error = %v", err)
	}
}

type flushRecorder struct {
	bytes.Buffer
	flushes int
}

func (w *flushRecorder) Flush() {
	w.flushes++
}

func TestStreamWriterFlushesToUnderlyingWriter(t *testing.T) {
	t.Parallel()

	var output flushRecorder
	workbook, err := NewStreamWriter(&output, Options{FlushEveryRows: 2})
	if err != nil {
		t.Fatal(err)
	}
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	for i := 0; i < 2; i++ {
		row, err := sheet.AddRow()
		if err != nil {
			t.Fatal(err)
		}
		if _, err := row.WriteInt(int64(i)); err != nil {
			t.Fatal(err)
		}
	}
	if err := sheet.CommitRows(); err != nil {
		t.Fatal(err)
	}
	if output.flushes == 0 {
		t.Fatal("underlying writer was not flushed")
	}
	if output.Len() == 0 {
		t.Fatal("no bytes were written before Commit")
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}
}

func TestNewWriterHTTPResponse(t *testing.T) {
	t.Parallel()

	recorder := httptest.NewRecorder()
	request := httptest.NewRequest(http.MethodGet, "/export.xlsx", nil)
	http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		w.Header().Set("Content-Disposition", `attachment; filename="export.xlsx"`)

		workbook, err := NewWriter(w, Options{Context: r.Context()})
		if err != nil {
			t.Error(err)
			return
		}
		sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
		row, err := sheet.AddRow()
		if err != nil {
			t.Error(err)
			return
		}
		if _, err := row.WriteString("buffered"); err != nil {
			t.Error(err)
			return
		}
		if err := sheet.CommitRows(); err != nil {
			t.Error(err)
			return
		}
		if err := workbook.Commit(); err != nil {
			t.Error(err)
		}
	}).ServeHTTP(recorder, request)

	response := recorder.Result()
	defer response.Body.Close()
	if response.StatusCode != http.StatusOK {
		t.Fatalf("status = %d, want 200", response.StatusCode)
	}
	if got := response.Header.Get("Content-Type"); got != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" {
		t.Fatalf("Content-Type = %q", got)
	}
	body, err := io.ReadAll(response.Body)
	if err != nil {
		t.Fatal(err)
	}
	if err := ValidateReaderAt(bytes.NewReader(body), int64(len(body))); err != nil {
		t.Fatalf("ValidateReaderAt() error = %v", err)
	}
}

func TestStreamWriterHTTPResponse(t *testing.T) {
	t.Parallel()

	recorder := httptest.NewRecorder()
	request := httptest.NewRequest(http.MethodGet, "/stream.xlsx", nil)
	http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		w.Header().Set("Content-Disposition", `attachment; filename="stream.xlsx"`)

		workbook, err := NewStreamWriter(w, Options{
			Context:        r.Context(),
			FlushEveryRows: 1,
			SharedStrings:  true,
		})
		if err != nil {
			t.Error(err)
			return
		}
		sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
		for i := 0; i < 3; i++ {
			row, err := sheet.AddRow()
			if err != nil {
				t.Error(err)
				return
			}
			if _, err := row.WriteString("streamed"); err != nil {
				t.Error(err)
				return
			}
			if _, err := row.WriteInt(int64(i)); err != nil {
				t.Error(err)
				return
			}
			if err := sheet.CommitRows(); err != nil {
				t.Error(err)
				return
			}
		}
		if err := workbook.Flush(); err != nil {
			t.Error(err)
			return
		}
		if err := workbook.Commit(); err != nil {
			t.Error(err)
		}
	}).ServeHTTP(recorder, request)

	if !recorder.Flushed {
		t.Fatal("response recorder was not flushed")
	}
	response := recorder.Result()
	defer response.Body.Close()
	body, err := io.ReadAll(response.Body)
	if err != nil {
		t.Fatal(err)
	}
	if err := ValidateReaderAt(bytes.NewReader(body), int64(len(body))); err != nil {
		t.Fatalf("ValidateReaderAt() error = %v", err)
	}
	reader, err := OpenReaderAt(bytes.NewReader(body), int64(len(body)))
	if err != nil {
		t.Fatal(err)
	}
	defer reader.Close()
	sheet, err := reader.OpenSheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	defer sheet.Close()
	rows := 0
	for sheet.Next() {
		rows++
	}
	if err := sheet.Err(); err != nil {
		t.Fatal(err)
	}
	if rows != 3 {
		t.Fatalf("read %d rows, want 3", rows)
	}
}

func TestStreamWriterDoesNotCreateTempDirs(t *testing.T) {
	t.Parallel()

	missing := filepath.Join(t.TempDir(), "missing")
	var buf bytes.Buffer
	workbook, err := NewStreamWriter(&buf, Options{TempDir: missing})
	if err != nil {
		t.Fatal(err)
	}
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	row, err := sheet.AddRow()
	if err != nil {
		t.Fatal(err)
	}
	if _, err := row.WriteString("x"); err != nil {
		t.Fatal(err)
	}
	if err := sheet.CommitRows(); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		t.Fatal(err)
	}
	if workbook.TempDir() != "" {
		t.Fatalf("TempDir() = %q, want empty", workbook.TempDir())
	}
}
