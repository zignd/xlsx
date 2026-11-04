package xlsx_test

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/zignd/xlsx"
)

func exampleWorkbook() (string, func()) {
	dir, err := os.MkdirTemp("", "xlsx-reader-example-*")
	if err != nil {
		panic(err)
	}
	path := filepath.Join(dir, "people.xlsx")
	workbook := xlsx.NewWorkbook(path)
	worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "People"})
	for _, person := range [][]interface{}{
		{"Ada", 36},
		{"Grace", 85},
	} {
		row, err := worksheet.AddRow()
		if err != nil {
			panic(err)
		}
		if err := row.Append(person...); err != nil {
			panic(err)
		}
	}
	if err := worksheet.CommitRows(); err != nil {
		panic(err)
	}
	if err := workbook.Commit(); err != nil {
		panic(err)
	}
	return path, func() { _ = os.RemoveAll(dir) }
}

func ExampleOpen() {
	path, cleanup := exampleWorkbook()
	defer cleanup()

	reader, err := xlsx.Open(path)
	if err != nil {
		panic(err)
	}
	defer reader.Close()

	sheet, err := reader.OpenSheet("People")
	if err != nil {
		panic(err)
	}
	defer sheet.Close()

	for sheet.Next() {
		name, _ := sheet.Row().CellAt(0)
		age, _ := sheet.Row().CellAt(1)
		nameValue, _ := name.String()
		ageValue, _ := age.Int64()
		fmt.Printf("%s is %d\n", nameValue, ageValue)
	}
	if err := sheet.Err(); err != nil {
		panic(err)
	}

	// Output:
	// Ada is 36
	// Grace is 85
}

func ExampleValidate() {
	path, cleanup := exampleWorkbook()
	defer cleanup()

	if err := xlsx.Validate(path); err != nil {
		panic(err)
	}
	fmt.Println("valid workbook")

	// Output: valid workbook
}

func ExampleNewWorkbookWithOptions() {
	dir, err := os.MkdirTemp("", "xlsx-options-example-*")
	if err != nil {
		panic(err)
	}
	defer os.RemoveAll(dir)

	workbook, err := xlsx.NewWorkbookWithOptions(filepath.Join(dir, "styled.xlsx"), xlsx.Options{
		TempDir:       dir,
		SharedStrings: true,
	})
	if err != nil {
		panic(err)
	}

	style, err := workbook.AddStyle(xlsx.Style{Bold: true, NumberFormat: "#,##0.00"})
	if err != nil {
		panic(err)
	}
	worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Data"})
	if err := worksheet.FreezePanes(1, 0); err != nil {
		panic(err)
	}
	row, err := worksheet.AddRow()
	if err != nil {
		panic(err)
	}
	amount, err := row.WriteFloat(1234.5)
	if err != nil {
		panic(err)
	}
	amount.SetStyle(style)
	amount.SetHyperlink("https://example.com")
	formula, err := row.AddCell()
	if err != nil {
		panic(err)
	}
	formula.SetFormula("=A1*2")
	if err := worksheet.CommitRows(); err != nil {
		panic(err)
	}
	if err := workbook.Commit(); err != nil {
		panic(err)
	}

	fmt.Printf("style ID: %d\n", style)
	// Output: style ID: 2
}

func ExampleWorksheet_MergeCells() {
	dir, err := os.MkdirTemp("", "xlsx-layout-example-*")
	if err != nil {
		panic(err)
	}
	defer os.RemoveAll(dir)

	workbook := xlsx.NewWorkbook(filepath.Join(dir, "layout.xlsx"))
	worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Data"})
	if err := worksheet.FreezePanes(1, 0); err != nil {
		panic(err)
	}
	if err := worksheet.SetAutoFilter("A1:B2"); err != nil {
		panic(err)
	}
	if err := worksheet.MergeCells("A3:B3"); err != nil {
		panic(err)
	}
	row, err := worksheet.AddRow()
	if err != nil {
		panic(err)
	}
	if err := row.Append("Name", "Amount"); err != nil {
		panic(err)
	}
	if err := worksheet.CommitRows(); err != nil {
		panic(err)
	}
	if err := workbook.Commit(); err != nil {
		panic(err)
	}
	fmt.Println("layout written")
	// Output: layout written
}
