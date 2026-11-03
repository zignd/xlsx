// Package xlsx lets you create XLSX files incrementally.
//
// Rows are buffered only until CommitRows is called and are then streamed to
// temporary files on disk. This keeps memory usage low even when writing very
// large spreadsheets.
//
// # Quick start
//
//	workbook := xlsx.NewWorkbook("report.xlsx")
//	worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Sales"})
//
//	row, err := worksheet.AddRow()
//	if err != nil {
//		log.Fatal(err)
//	}
//	cell, err := row.AddCell()
//	if err != nil {
//		log.Fatal(err)
//	}
//	cell.Value = "Revenue"
//
//	if err := worksheet.CommitRows(); err != nil {
//		log.Fatal(err)
//	}
//	if err := worksheet.Commit(); err != nil {
//		log.Fatal(err)
//	}
//	if err := workbook.Commit(); err != nil {
//		log.Fatal(err)
//	}
//
// # Values
//
// [Cell.Value] accepts strings, booleans, all integer and unsigned integer
// types, floats, [time.Time], []byte and [database/sql/driver.Valuer]
// implementations (such as the sql.Null* types). Unsupported values are
// reported when rows are committed.
//
// # Streaming
//
// Call [Worksheet.CommitRows] after adding a batch of rows to release the rows
// from memory. The worksheet keeps a buffered writer open, so repeated batches
// are cheap to append.
package xlsx
