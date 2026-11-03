package xlsx

import (
	"fmt"
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
