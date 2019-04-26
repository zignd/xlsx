package xlsx

import (
	"fmt"
	"testing"
	"time"
)

func Test_Usage(t *testing.T) {
	workbook := NewWorkbook("./spreadsheet-1.xlsx")
	worksheet := workbook.AddWorksheet(&WorksheetOptions{
		Name: "Data",
	})

	for i := 0; i < 3; i++ {
		row, _ := worksheet.AddRow()

		cell, _ := row.AddCell()
		cell.Value = "test 1"

		cell, _ = row.AddCell()
		cell.Value = 2

		cell, _ = row.AddCell()
		cell.Value = time.Now()

		cell, _ = row.AddCell()
		cell.Value = "test 4"

		cell, _ = row.AddCell()
		cell.Value = "test 5"

		cell, _ = row.AddCell()
		cell.Value = "test 6"

		cell, _ = row.AddCell()
		cell.Value = "test 7"

		cell, _ = row.AddCell()
		cell.Value = "test 8"

		cell, _ = row.AddCell()
		cell.Value = "test 9"

		cell, _ = row.AddCell()
		cell.Value = "test 10"

		err := worksheet.CommitRows()
		if i%50000 == 0 {
			fmt.Println(i)
		}
		if err != nil {
			t.Errorf("failed to commit rows: %v", err)
			return
		}
	}
	err := worksheet.Commit()
	if err != nil {
		t.Errorf("failed to commit worksheet: %v", err)
		return
	}
	err = workbook.Commit()
	if err != nil {
		t.Errorf("failed to commit workbook: %v", err)
	}

	fmt.Println("workbook.tempRootDir", workbook.tempRootDir)
}

func Test_Usage_AddCellWithKey(t *testing.T) {
	workbook := NewWorkbook("./spreadsheet-2.xlsx")
	worksheet := workbook.AddWorksheet(&WorksheetOptions{
		Name: "Data",
	})

	worksheet.DefineColumns([]*WorksheetColumn{
		&WorksheetColumn{Key: "sup1", Value: "sup 1"},
		&WorksheetColumn{Key: "sup2", Value: "sup 2"},
		&WorksheetColumn{Key: "sup3", Value: "sup 3"},
	})

	for i := 0; i < 3; i++ {
		row, _ := worksheet.AddRow()

		cell, err := row.AddCellWithKey("sup1")
		if err != nil {
			t.Errorf("failed to add cell: %v", err)
		}
		cell.Value = "test 1"

		if i%2 == 0 {
			cell, err = row.AddCellWithKey("sup2")
			if err != nil {
				t.Errorf("failed to add cell: %v", err)
				return
			}
			cell.Value = 2
		}

		cell, err = row.AddCellWithKey("sup3")
		if err != nil {
			t.Errorf("failed to add cell: %v", err)
			return
		}
		cell.Value = time.Now()

		err = worksheet.CommitRows()
		if err != nil {
			t.Errorf("failed to commit rows: %v", err)
			return
		}

		if i%50000 == 0 {
			fmt.Println(i)
		}
	}
	err := worksheet.Commit()
	if err != nil {
		t.Errorf("failed to commit worksheet: %v", err)
		return
	}
	err = workbook.Commit()
	if err != nil {
		t.Errorf("failed to commit workbook: %v", err)
	}

	fmt.Println("workbook.tempRootDir", workbook.tempRootDir)
}
