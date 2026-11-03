package xlsx_test

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/zignd/xlsx"
)

func ExampleNewWorkbook() {
	dir, err := os.MkdirTemp("", "xlsx-example-*")
	if err != nil {
		panic(err)
	}
	defer os.RemoveAll(dir)

	workbook := xlsx.NewWorkbook(filepath.Join(dir, "report.xlsx"))
	worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Sales"})

	for i := 0; i < 3; i++ {
		row, err := worksheet.AddRow()
		if err != nil {
			panic(err)
		}

		product, err := row.AddCell()
		if err != nil {
			panic(err)
		}
		product.Value = fmt.Sprintf("Product %d", i+1)

		price, err := row.AddCell()
		if err != nil {
			panic(err)
		}
		price.Value = 9.99 * float64(i+1)

		if err := worksheet.CommitRows(); err != nil {
			panic(err)
		}
	}

	if err := workbook.Commit(); err != nil {
		panic(err)
	}

	fmt.Println("created report.xlsx with 3 rows")
	// Output: created report.xlsx with 3 rows
}

func ExampleWorksheet_DefineColumns() {
	dir, err := os.MkdirTemp("", "xlsx-example-*")
	if err != nil {
		panic(err)
	}
	defer os.RemoveAll(dir)

	workbook := xlsx.NewWorkbook(filepath.Join(dir, "report.xlsx"))
	worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Sales"})

	err = worksheet.DefineColumns([]*xlsx.WorksheetColumn{
		{Key: "product", Value: "Product", Width: 30},
		{Key: "amount", Value: "Amount"},
	})
	if err != nil {
		panic(err)
	}

	row, err := worksheet.AddRow()
	if err != nil {
		panic(err)
	}
	product, err := row.AddCellWithKey("product")
	if err != nil {
		panic(err)
	}
	product.Value = "Keyboard"
	amount, err := row.AddCellWithKey("amount")
	if err != nil {
		panic(err)
	}
	amount.Value = 49.9

	if err := worksheet.CommitRows(); err != nil {
		panic(err)
	}
	if err := workbook.Commit(); err != nil {
		panic(err)
	}

	fmt.Println("created a keyed worksheet")
	// Output: created a keyed worksheet
}
