package xlsx

import (
	"database/sql"
	"database/sql/driver"
	"errors"
	"fmt"
	"math"
	"path/filepath"
	"testing"
)

type failingValuer struct{}

func (failingValuer) Value() (driver.Value, error) {
	return nil, errors.New("valuer failed")
}

type customString string
type customBool bool
type customInt int
type customFloat float64

func TestNamedTypes(t *testing.T) {
	t.Parallel()

	wb, ws := newTestWorksheet(t)
	row, err := ws.AddRow()
	if err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	for _, value := range []interface{}{
		customString("text"),
		customBool(true),
		customInt(-3),
		customFloat(1.25),
		sql.RawBytes("raw bytes"),
	} {
		cell, err := row.AddCell()
		if err != nil {
			t.Fatalf("AddCell() error = %v", err)
		}
		cell.Value = value
	}
	if err := ws.CommitRows(); err != nil {
		t.Fatalf("CommitRows() error = %v", err)
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}

func newTestWorksheet(t *testing.T) (*Workbook, *Worksheet) {
	t.Helper()
	wb := NewWorkbook(filepath.Join(t.TempDir(), "workbook.xlsx"))
	return wb, wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
}

func TestDefineColumns(t *testing.T) {
	t.Parallel()

	wb, ws := newTestWorksheet(t)
	err := ws.DefineColumns([]*WorksheetColumn{
		{Key: "name", Value: "Name", Width: 20},
		{Key: "age", Value: "Age"},
	})
	if err != nil {
		t.Fatalf("DefineColumns() error = %v", err)
	}
	if got, want := ws.rowsCount, 1; got != want {
		t.Fatalf("rowsCount = %d, want %d", got, want)
	}

	row, err := ws.AddRow()
	if err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	cell, err := row.AddCellWithKey("name")
	if err != nil {
		t.Fatalf("AddCellWithKey() error = %v", err)
	}
	cell.Value = "Ada"

	if err := ws.CommitRows(); err != nil {
		t.Fatalf("CommitRows() error = %v", err)
	}
	if err := ws.Commit(); err != nil {
		t.Fatalf("Commit() error = %v", err)
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}

func TestDefineColumnsValidation(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name    string
		columns []*WorksheetColumn
	}{
		{"empty", nil},
		{"nil column", []*WorksheetColumn{nil}},
		{"empty key", []*WorksheetColumn{{Key: ""}}},
		{"duplicate key", []*WorksheetColumn{{Key: "a"}, {Key: "a"}}},
		{"negative width", []*WorksheetColumn{{Key: "a", Width: -1}}},
		{"excessive width", []*WorksheetColumn{{Key: "a", Width: 256}}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			wb, ws := newTestWorksheet(t)
			if err := ws.DefineColumns(tt.columns); !errors.Is(err, ErrInvalidColumn) {
				t.Fatalf("DefineColumns() error = %v, want ErrInvalidColumn", err)
			}
			if err := wb.Cleanup(); err != nil {
				t.Fatalf("Cleanup() error = %v", err)
			}
		})
	}
}

func TestDefineColumnsAfterRowsWereCommitted(t *testing.T) {
	t.Parallel()

	wb, ws := newTestWorksheet(t)
	row, err := ws.AddRow()
	if err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}
	if _, err := row.AddCell(); err != nil {
		t.Fatalf("AddCell() error = %v", err)
	}
	if err := ws.CommitRows(); err != nil {
		t.Fatalf("CommitRows() error = %v", err)
	}

	if err := ws.DefineColumns([]*WorksheetColumn{{Key: "a"}}); err == nil {
		t.Fatal("DefineColumns() returned nil after rows were committed")
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}

func TestDefineColumnsWithPendingRows(t *testing.T) {
	t.Parallel()

	wb, ws := newTestWorksheet(t)
	if _, err := ws.AddRow(); err != nil {
		t.Fatalf("AddRow() error = %v", err)
	}

	if err := ws.DefineColumns([]*WorksheetColumn{{Key: "a"}}); err == nil {
		t.Fatal("DefineColumns() returned nil while rows were pending")
	}
	if err := wb.Cleanup(); err != nil {
		t.Fatalf("Cleanup() error = %v", err)
	}
}

func TestAddCellValidation(t *testing.T) {
	t.Parallel()

	t.Run("columns defined", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		if err := ws.DefineColumns([]*WorksheetColumn{{Key: "a"}}); err != nil {
			t.Fatalf("DefineColumns() error = %v", err)
		}
		row, err := ws.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		if _, err := row.AddCell(); !errors.Is(err, ErrColumnsDefined) {
			t.Fatalf("AddCell() error = %v, want ErrColumnsDefined", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("committed row", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		row, err := ws.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		if _, err := row.AddCell(); err != nil {
			t.Fatalf("AddCell() error = %v", err)
		}
		if err := ws.CommitRows(); err != nil {
			t.Fatalf("CommitRows() error = %v", err)
		}
		if _, err := row.AddCell(); err == nil {
			t.Fatal("AddCell() returned nil for a committed row")
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})
}

func TestAddCellWithKeyValidation(t *testing.T) {
	t.Parallel()

	t.Run("no columns", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		row, err := ws.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		if _, err := row.AddCellWithKey("a"); !errors.Is(err, ErrNoColumns) {
			t.Fatalf("AddCellWithKey() error = %v, want ErrNoColumns", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("unknown key", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		if err := ws.DefineColumns([]*WorksheetColumn{{Key: "a"}}); err != nil {
			t.Fatalf("DefineColumns() error = %v", err)
		}
		row, err := ws.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		if _, err := row.AddCellWithKey("missing"); err == nil {
			t.Fatal("AddCellWithKey() returned nil for an unknown key")
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("duplicate cell", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		if err := ws.DefineColumns([]*WorksheetColumn{{Key: "a"}}); err != nil {
			t.Fatalf("DefineColumns() error = %v", err)
		}
		row, err := ws.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		if _, err := row.AddCellWithKey("a"); err != nil {
			t.Fatalf("first AddCellWithKey() error = %v", err)
		}
		if _, err := row.AddCellWithKey("a"); err == nil {
			t.Fatal("second AddCellWithKey() returned nil")
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})
}

func TestWorksheetLimits(t *testing.T) {
	t.Parallel()

	t.Run("too many columns", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		columns := make([]*WorksheetColumn, maxWorksheetColumns+1)
		for i := range columns {
			columns[i] = &WorksheetColumn{Key: fmt.Sprintf("column-%d", i)}
		}
		if err := ws.DefineColumns(columns); !errors.Is(err, ErrInvalidColumn) {
			t.Fatalf("DefineColumns() error = %v, want ErrInvalidColumn", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("too many cells", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		row, err := ws.AddRow()
		if err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		for i := 0; i < maxWorksheetColumns; i++ {
			if _, err := row.AddCell(); err != nil {
				t.Fatalf("AddCell() at %d returned error: %v", i, err)
			}
		}
		if _, err := row.AddCell(); err == nil {
			t.Fatal("AddCell() returned nil after the column limit")
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("too many rows", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		ws.rowsCount = maxWorksheetRows
		if _, err := ws.AddRow(); err == nil {
			t.Fatal("AddRow() returned nil after the row limit")
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})
}

func TestCommitRowsAndCommitValidation(t *testing.T) {
	t.Parallel()

	t.Run("no pending rows", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		if err := ws.CommitRows(); !errors.Is(err, ErrNoRows) {
			t.Fatalf("CommitRows() error = %v, want ErrNoRows", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("pending rows", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		if _, err := ws.AddRow(); err != nil {
			t.Fatalf("AddRow() error = %v", err)
		}
		if err := ws.Commit(); !errors.Is(err, ErrPendingRows) {
			t.Fatalf("Commit() error = %v, want ErrPendingRows", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("committed worksheet", func(t *testing.T) {
		t.Parallel()
		wb, ws := newTestWorksheet(t)
		if err := ws.Commit(); err != nil {
			t.Fatalf("Commit() error = %v", err)
		}
		if err := ws.Commit(); !errors.Is(err, ErrWorksheetCommitted) {
			t.Fatalf("second Commit() error = %v, want ErrWorksheetCommitted", err)
		}
		if _, err := ws.AddRow(); !errors.Is(err, ErrWorksheetCommitted) {
			t.Fatalf("AddRow() error = %v, want ErrWorksheetCommitted", err)
		}
		if err := ws.CommitRows(); !errors.Is(err, ErrWorksheetCommitted) {
			t.Fatalf("CommitRows() error = %v, want ErrWorksheetCommitted", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})
}

func TestUnsupportedCellValues(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name  string
		value interface{}
	}{
		{"struct", struct{ A int }{1}},
		{"map", map[string]int{"a": 1}},
		{"channel", make(chan int)},
		{"nan", math.NaN()},
		{"infinite", math.Inf(1)},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
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
			cell.Value = tt.value
			if err := ws.CommitRows(); !errors.Is(err, ErrUnsupportedType) {
				t.Fatalf("CommitRows() error = %v, want ErrUnsupportedType", err)
			}
			if err := wb.Cleanup(); err != nil {
				t.Fatalf("Cleanup() error = %v", err)
			}
		})
	}
}

func TestDriverValuer(t *testing.T) {
	t.Parallel()

	t.Run("null string", func(t *testing.T) {
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
		cell.Value = sql.NullString{String: "hello", Valid: true}
		if err := ws.CommitRows(); err != nil {
			t.Fatalf("CommitRows() error = %v", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("invalid null string", func(t *testing.T) {
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
		cell.Value = sql.NullString{}
		if err := ws.CommitRows(); err != nil {
			t.Fatalf("CommitRows() error = %v", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("typed nil valuer", func(t *testing.T) {
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
		var value *sql.NullString
		cell.Value = value
		if err := ws.CommitRows(); err != nil {
			t.Fatalf("CommitRows() error = %v", err)
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})

	t.Run("failing valuer", func(t *testing.T) {
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
		cell.Value = failingValuer{}
		if err := ws.CommitRows(); err == nil {
			t.Fatal("CommitRows() returned nil for a failing driver.Valuer")
		}
		if err := wb.Cleanup(); err != nil {
			t.Fatalf("Cleanup() error = %v", err)
		}
	})
}
