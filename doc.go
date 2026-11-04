// Package xlsx lets you create and read XLSX files incrementally.
//
// The writer buffers rows only until CommitRows is called and then streams
// them to temporary files on disk. The reader streams worksheet rows with
// xml.Decoder and keeps only one row at a time in memory. Both directions are
// designed for large workbooks.
//
// # Writing
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
//	cell.SetString("Revenue")
//
//	if err := worksheet.CommitRows(); err != nil {
//		log.Fatal(err)
//	}
//	if err := workbook.Commit(); err != nil {
//		log.Fatal(err)
//	}
//
// # Reading
//
//	reader, err := xlsx.Open("report.xlsx")
//	if err != nil {
//		log.Fatal(err)
//	}
//	defer reader.Close()
//
//	sheet, err := reader.OpenSheet("Sales")
//	if err != nil {
//		log.Fatal(err)
//	}
//	defer sheet.Close()
//
//	for sheet.Next() {
//		cell, ok := sheet.Row().CellAt(0)
//		if !ok {
//			continue
//		}
//		value, err := cell.String()
//		if err != nil {
//			log.Fatal(err)
//		}
//		fmt.Println(value)
//	}
//	if err := sheet.Err(); err != nil {
//		log.Fatal(err)
//	}
//
// # Options and temporary files
//
// [NewWorkbookWithOptions] accepts a temp directory, buffer size, compression,
// file modes, context cancellation, shared strings and safety limits. The
// library creates a private subdirectory inside the configured temp directory
// and never deletes the parent directory.
//
// [NewWriter] writes the final package to an io.Writer, which is useful for
// HTTP responses and object storage.
//
// # Values
//
// [Cell.Value] accepts strings, booleans, all integer and unsigned integer
// types, floats, [time.Time], []byte and [database/sql/driver.Valuer]
// implementations. Use [Cell.SetFormula] for formulas, [Cell.SetHyperlink] for
// links and [Cell.SetStyle] with an ID returned by [Workbook.AddStyle].
//
// # Validation
//
// [Validate] checks the package structure, content types, relationships,
// worksheet references, styles and shared strings without loading the whole
// workbook into memory.
package xlsx
