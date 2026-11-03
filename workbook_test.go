package xlsx

import (
	"errors"
	"os"
	"path/filepath"
	"testing"
)

func TestAddWorksheet(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Sheet 1"})

	if got, want := ws.Name(), "Sheet 1"; got != want {
		t.Fatalf("Name() = %q, want %q", got, want)
	}
	if got, want := ws.FileName(), "sheet1.xml"; got != want {
		t.Fatalf("FileName() = %q, want %q", got, want)
	}
	if got, want := ws.ID(), 1; got != want {
		t.Fatalf("ID() = %d, want %d", got, want)
	}

	second := wb.AddWorksheet(&WorksheetOptions{Name: "Sheet 2"})
	if got, want := second.FileName(), "sheet2.xml"; got != want {
		t.Fatalf("second FileName() = %q, want %q", got, want)
	}
	if got, want := second.ID(), 2; got != want {
		t.Fatalf("second ID() = %d, want %d", got, want)
	}
}

func TestWorksheetsReturnsCopy(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	wb.AddWorksheet(&WorksheetOptions{Name: "Sheet 1"})

	worksheets := wb.Worksheets()
	worksheets[0] = nil

	if wb.Worksheets()[0] == nil {
		t.Fatal("Worksheets() exposed the internal slice")
	}
}

func TestCreateTempDirs(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	if err := wb.createTempDirs(); err != nil {
		t.Fatalf("createTempDirs() error = %v", err)
	}

	if !wb.tempDirsCreated {
		t.Fatal("tempDirsCreated = false, want true")
	}
	for _, dir := range []string{
		wb.tempRootDir,
		wb.tempDirs.rels,
		wb.tempDirs.xl,
		wb.tempDirs.xlRels,
		wb.tempDirs.xlWorksheets,
	} {
		info, err := os.Stat(dir)
		if err != nil {
			t.Fatalf("stat(%q) error = %v", dir, err)
		}
		if !info.IsDir() {
			t.Fatalf("%q is not a directory", dir)
		}
	}

	if err := wb.createTempDirs(); err == nil {
		t.Fatal("createTempDirs() second call returned nil, want error")
	}

	root := wb.tempRootDir
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
	if _, err := os.Stat(root); !os.IsNotExist(err) {
		t.Fatalf("temporary root still exists, stat error = %v", err)
	}
	if wb.tempDirsCreated {
		t.Fatal("tempDirsCreated = true after Cleanup")
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("second Cleanup() error = %v", err)
	}
}

func TestCommitRequiresWorksheet(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	if err := wb.Commit(); !errors.Is(err, ErrNoWorksheets) {
		t.Fatalf("Commit() error = %v, want ErrNoWorksheets", err)
	}
}

func TestCommitRequiresFilePath(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook("")
	wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := wb.Commit(); err == nil {
		t.Fatal("Commit() returned nil for an empty file path")
	}
}

func TestCommitInvalidOutputPath(t *testing.T) {
	t.Parallel()

	dir := t.TempDir()
	wb := NewWorkbook(dir)
	wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := wb.Commit(); err == nil {
		t.Fatal("Commit() returned nil when the output path is a directory")
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}

func TestCommitRejectsPendingRows(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if _, err := ws.AddRow(); err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}

	if err := wb.Commit(); !errors.Is(err, ErrPendingRows) {
		t.Fatalf("Commit() error = %v, want ErrPendingRows", err)
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}

func TestCommitTwice(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "workbook.xlsx")
	wb := NewWorkbook(path)
	wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := wb.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	if err := wb.Commit(); !errors.Is(err, ErrWorkbookCommitted) {
		t.Fatalf("second Commit() error = %v, want ErrWorkbookCommitted", err)
	}
}

func TestCommitAutoFinalizesEmptyWorksheet(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "workbook.xlsx")
	wb := NewWorkbook(path)
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Empty"})

	if err := wb.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	if !ws.committed {
		t.Fatal("worksheet was not finalized automatically")
	}
	if _, err := os.Stat(path); err != nil {
		t.Fatalf("output file does not exist: %v", err)
	}
}

func TestInvalidWorksheetNames(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name      string
		sheetName string
		want      error
	}{
		{"empty", "", ErrInvalidSheetName},
		{"too long", "12345678901234567890123456789012", ErrInvalidSheetName},
		{"invalid character", "Sales/2024", ErrInvalidSheetName},
		{"control character", "Sales\n2024", ErrInvalidSheetName},
		{"leading apostrophe", "'Sales", ErrInvalidSheetName},
		{"reserved", "history", ErrInvalidSheetName},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
			ws := wb.AddWorksheet(&WorksheetOptions{Name: tt.sheetName})
			if err := ws.Commit(); !errors.Is(err, tt.want) {
				t.Fatalf("Commit() error = %v, want %v", err, tt.want)
			}
		})
	}
}

func TestDuplicateWorksheetNames(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	if err := wb.AddWorksheet(&WorksheetOptions{Name: "Data"}).Commit(); err != nil {
		t.Fatalf("first worksheet Commit() error = %v", err)
	}
	duplicate := wb.AddWorksheet(&WorksheetOptions{Name: "data"})
	if err := duplicate.Commit(); !errors.Is(err, ErrDuplicateSheetName) {
		t.Fatalf("Commit() error = %v, want ErrDuplicateSheetName", err)
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}

func TestAddWorksheetAfterCommit(t *testing.T) {
	t.Parallel()

	path := filepath.Join(t.TempDir(), "workbook.xlsx")
	wb := NewWorkbook(path)
	wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if err := wb.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}

	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Later"})
	if err := ws.Commit(); !errors.Is(err, ErrWorkbookCommitted) {
		t.Fatalf("Commit() error = %v, want ErrWorkbookCommitted", err)
	}
}

func TestCommitAfterCleanup(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if _, err := ws.AddRow(); err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	if err := ws.CommitRows(); err != nil {
		t.Fatalf("CommitRows() error = %v", err)
	}
	if err := ws.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
	if err := wb.Commit(); err == nil {
		t.Fatal("Commit() returned nil after Cleanup")
	}
}

func TestHasPendingWorksheets(t *testing.T) {
	t.Parallel()

	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	if wb.HasPendingWorksheets() {
		t.Fatal("HasPendingWorksheets() = true for an empty workbook")
	}

	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	if !wb.HasPendingWorksheets() {
		t.Fatal("HasPendingWorksheets() = false after AddWorksheet")
	}

	if err := ws.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	if wb.HasPendingWorksheets() {
		t.Fatal("HasPendingWorksheets() = true after the worksheet was committed")
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}
