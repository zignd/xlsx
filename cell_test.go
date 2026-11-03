package xlsx

import (
	"path/filepath"
	"testing"
)

func TestCellAccessors(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := ws.DefineColumns([]*WorksheetColumn{{Key: "first"}, {Key: "second"}}); err != nil {
		t.Fatalf("DefineColumns() error = %v", err)
	}

	row, err := ws.AddRow()
	if err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	if got, want := row.Index(), 1; got != want {
		t.Fatalf("row.Index() = %d, want %d", got, want)
	}

	first, err := row.AddCellWithKey("first")
	if err != nil {
		t.Fatalf("AddCellWithKey() error = %v", err)
	}
	second, err := row.AddCellWithKey("second")
	if err != nil {
		t.Fatalf("AddCellWithKey() error = %v", err)
	}

	if got, want := first.Identifier(), "A2"; got != want {
		t.Errorf("first.Identifier() = %q, want %q", got, want)
	}
	if got, want := first.Index(), 0; got != want {
		t.Errorf("first.Index() = %d, want %d", got, want)
	}
	if got, want := second.Identifier(), "B2"; got != want {
		t.Errorf("second.Identifier() = %q, want %q", got, want)
	}
	if got, want := second.Index(), 1; got != want {
		t.Errorf("second.Index() = %d, want %d", got, want)
	}
	if first.Row() != row {
		t.Error("Cell.Row() did not return the owning row")
	}
	if first.Key != "first" || second.Key != "second" {
		t.Errorf("unexpected cell keys: %q, %q", first.Key, second.Key)
	}

	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}

func TestCellAccessorsWithoutColumns(t *testing.T) {
	t.Parallel()

	wb, ws := newTestWorksheet(t)
	row, err := ws.AddRow()
	if err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	cell, err := row.AddCell()
	if err != nil {
		t.Fatalf("AddCell() error = %v", err)
	}
	if got, want := cell.Identifier(), "A1"; got != want {
		t.Errorf("Identifier() = %q, want %q", got, want)
	}
	if got, want := cell.Index(), 0; got != want {
		t.Errorf("Index() = %d, want %d", got, want)
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}
