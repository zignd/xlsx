package xlsx

import (
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"testing"
	"time"
)

func BenchmarkNumericToLetters(b *testing.B) {
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		_ = numericToLetters(i % 16384)
	}
}

func BenchmarkTimeToExcelTime(b *testing.B) {
	value := time.Date(2024, 1, 2, 3, 4, 5, 123456789, time.UTC)
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		_ = timeToExcelTime(value)
	}
}

func BenchmarkXMLEscape(b *testing.B) {
	value := `a & b < c > d "quoted" and some more text to escape`
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		_ = xmlEscape(value)
	}
}

func benchmarkRows(b *testing.B, worksheet *Worksheet, count int) {
	b.Helper()

	for i := 0; i < count; i++ {
		row, err := worksheet.AddRow()
		if err != nil {
			b.Fatalf("AddRow() error = %v", err)
		}
		for j := 0; j < 5; j++ {
			cell, err := row.AddCell()
			if err != nil {
				b.Fatalf("AddCell() error = %v", err)
			}
			cell.Value = fmt.Sprintf("value-%d-%d", i, j)
		}
	}
}

func BenchmarkAddRows(b *testing.B) {
	b.ReportAllocs()

	wb := NewWorkbook(filepath.Join(b.TempDir(), "workbook.xlsx"))
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		row, err := ws.AddRow()
		if err != nil {
			b.Fatalf("AddRow() error = %v", err)
		}
		for j := 0; j < 5; j++ {
			cell, err := row.AddCell()
			if err != nil {
				b.Fatalf("AddCell() error = %v", err)
			}
			cell.Value = i
		}
	}
}

func BenchmarkCommitRows(b *testing.B) {
	for _, count := range []int{100, 1000} {
		b.Run(fmt.Sprintf("rows-%d", count), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				path := filepath.Join(b.TempDir(), fmt.Sprintf("workbook-%d.xlsx", i))
				wb := NewWorkbook(path)
				ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
				benchmarkRows(b, ws, count)
				b.StartTimer()

				if err := ws.CommitRows(); err != nil {
					b.Fatalf("CommitRows() error = %v", err)
				}

				b.StopTimer()
				if err := wb.Cleanup(); err != nil {
					b.Fatalf("Cleanup() error = %v", err)
				}
				b.StartTimer()
			}
		})
	}
}

func BenchmarkCommitWorkbook(b *testing.B) {
	for _, count := range []int{100, 1000} {
		b.Run(fmt.Sprintf("rows-%d", count), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				path := filepath.Join(b.TempDir(), fmt.Sprintf("workbook-%d.xlsx", i))
				wb := NewWorkbook(path)
				ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
				benchmarkRows(b, ws, count)
				b.StartTimer()

				if err := ws.CommitRows(); err != nil {
					b.Fatalf("CommitRows() error = %v", err)
				}
				if err := wb.Commit(); err != nil {
					b.Fatalf("Commit() error = %v", err)
				}
			}
		})
	}
}

func BenchmarkSharedStrings(b *testing.B) {
	for _, shared := range []bool{false, true} {
		b.Run(fmt.Sprintf("shared-%t", shared), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				path := filepath.Join(b.TempDir(), fmt.Sprintf("shared-%d.xlsx", i))
				wb, err := NewWorkbookWithOptions(path, Options{SharedStrings: shared})
				if err != nil {
					b.Fatal(err)
				}
				ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
				for rowIndex := 0; rowIndex < 100; rowIndex++ {
					row, err := ws.AddRow()
					if err != nil {
						b.Fatal(err)
					}
					for column := 0; column < 5; column++ {
						if _, err := row.WriteString("repeated value"); err != nil {
							b.Fatal(err)
						}
					}
				}
				b.StartTimer()
				if err := ws.CommitRows(); err != nil {
					b.Fatal(err)
				}
				if err := wb.Commit(); err != nil {
					b.Fatal(err)
				}
			}
		})
	}
}

func BenchmarkTypedWrites(b *testing.B) {
	b.ReportAllocs()
	wb := NewWorkbook(filepath.Join(b.TempDir(), "typed.xlsx"))
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		row, err := ws.AddRow()
		if err != nil {
			b.Fatal(err)
		}
		if _, err := row.WriteString("name"); err != nil {
			b.Fatal(err)
		}
		if _, err := row.WriteInt(int64(i)); err != nil {
			b.Fatal(err)
		}
		if _, err := row.WriteFloat(1.25); err != nil {
			b.Fatal(err)
		}
		if _, err := row.WriteBool(true); err != nil {
			b.Fatal(err)
		}
	}
}

func benchmarkWorkbook(b *testing.B, rows int) string {
	b.Helper()
	path := filepath.Join(b.TempDir(), "reader.xlsx")
	wb := NewWorkbook(path)
	ws := wb.AddWorksheet(&WorksheetOptions{Name: "Data"})
	for i := 0; i < rows; i++ {
		row, err := ws.AddRow()
		if err != nil {
			b.Fatal(err)
		}
		if _, err := row.WriteString("value"); err != nil {
			b.Fatal(err)
		}
		if _, err := row.WriteInt(int64(i)); err != nil {
			b.Fatal(err)
		}
	}
	if err := ws.CommitRows(); err != nil {
		b.Fatal(err)
	}
	if err := wb.Commit(); err != nil {
		b.Fatal(err)
	}
	return path
}

func BenchmarkReader(b *testing.B) {
	for _, rows := range []int{100, 1_000} {
		b.Run(fmt.Sprintf("rows-%d", rows), func(b *testing.B) {
			path := benchmarkWorkbook(b, rows)
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				reader, err := Open(path)
				if err != nil {
					b.Fatal(err)
				}
				sheet, err := reader.OpenSheet("Data")
				if err != nil {
					b.Fatal(err)
				}
				count := 0
				for sheet.Next() {
					count++
				}
				if err := sheet.Err(); err != nil {
					b.Fatal(err)
				}
				if count != rows {
					b.Fatalf("read %d rows, want %d", count, rows)
				}
				_ = sheet.Close()
				_ = reader.Close()
			}
		})
	}
}

func BenchmarkValidate(b *testing.B) {
	path := benchmarkWorkbook(b, 1_000)
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if err := Validate(path); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkStyledFormulas(b *testing.B) {
	path := filepath.Join(b.TempDir(), "styled-formulas.xlsx")
	workbook, err := NewWorkbookWithOptions(path, Options{SharedStrings: true})
	if err != nil {
		b.Fatal(err)
	}
	styleID, err := workbook.AddStyle(Style{Bold: true, NumberFormat: "#,##0.00", FillColor: "FFF2CC"})
	if err != nil {
		b.Fatal(err)
	}
	sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		row, err := sheet.AddRow()
		if err != nil {
			b.Fatal(err)
		}
		cell, err := row.WriteString("styled")
		if err != nil {
			b.Fatal(err)
		}
		cell.SetStyle(styleID)
		formula, err := row.AddCell()
		if err != nil {
			b.Fatal(err)
		}
		formula.SetFormula("A1&\"x\"")
	}
	b.StopTimer()
	if err := sheet.CommitRows(); err != nil {
		b.Fatal(err)
	}
	if err := workbook.Commit(); err != nil {
		b.Fatal(err)
	}
}

func BenchmarkWriteModes(b *testing.B) {
	for _, mode := range []struct {
		name   string
		stream bool
	}{
		{name: "buffered", stream: false},
		{name: "streaming", stream: true},
	} {
		b.Run(mode.name, func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				var (
					workbook *Workbook
					err      error
				)
				if mode.stream {
					workbook, err = NewStreamWriter(io.Discard, Options{FlushEveryRows: 250})
				} else {
					workbook, err = NewWriter(io.Discard, Options{})
				}
				if err != nil {
					b.Fatal(err)
				}
				sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
				b.StartTimer()

				for rowIndex := 0; rowIndex < 1_000; rowIndex++ {
					row, err := sheet.AddRow()
					if err != nil {
						b.Fatal(err)
					}
					if err := row.Append(rowIndex, "value", 1.25); err != nil {
						b.Fatal(err)
					}
				}
				if err := sheet.CommitRows(); err != nil {
					b.Fatal(err)
				}
				if err := workbook.Commit(); err != nil {
					b.Fatal(err)
				}
			}
		})
	}
}

func BenchmarkHTTPExport(b *testing.B) {
	for _, stream := range []bool{false, true} {
		name := "buffered"
		if stream {
			name = "streaming"
		}
		b.Run(name, func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				recorder := httptest.NewRecorder()
				request := httptest.NewRequest(http.MethodGet, "/export.xlsx", nil)
				handler := func(w http.ResponseWriter, r *http.Request) {
					w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
					var (
						workbook *Workbook
						err      error
					)
					if stream {
						workbook, err = NewStreamWriter(w, Options{Context: r.Context(), FlushEveryRows: 250})
					} else {
						workbook, err = NewWriter(w, Options{Context: r.Context()})
					}
					if err != nil {
						b.Fatal(err)
					}
					sheet := workbook.AddWorksheet(&WorksheetOptions{Name: "Data"})
					for rowIndex := 0; rowIndex < 1_000; rowIndex++ {
						row, err := sheet.AddRow()
						if err != nil {
							b.Fatal(err)
						}
						if err := row.Append(rowIndex, "value", 1.25); err != nil {
							b.Fatal(err)
						}
					}
					if err := sheet.CommitRows(); err != nil {
						b.Fatal(err)
					}
					if err := workbook.Commit(); err != nil {
						b.Fatal(err)
					}
				}
				http.HandlerFunc(handler).ServeHTTP(recorder, request)
			}
		})
	}
}
