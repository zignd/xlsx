package xlsx

import (
	"os"
	"testing"
)

func Test_Workbook_AddWorksheet_ShouldProperlyAddNewWorksheet_WhenGivenValidArguments(t *testing.T) {
	wb := NewWorkbook("./spreadsheet.xlsx")
	wsName := "Sheet 1"
	ws := wb.AddWorksheet(&WorksheetOptions{
		Name: wsName,
	})
	if ws.name != wsName {
		t.Errorf("worksheet name differs from the one given to AddWorksheet, expected \"%s\", found \"%s\"", wsName, ws.name)
		return
	}
	if ws.fileName != "sheet1.xml" {
		t.Errorf("worksheet file name is wrong, expected \"sheet1.xml\", found \"%s\"", ws.fileName)
		return
	}
}

func Test_Workbook_createTempDirs_ShouldProperlyCreateTemporaryDirectories(t *testing.T) {
	wb := NewWorkbook("./spreadsheet.xlsx")

	err := wb.createTempDirs()
	if err != nil {
		t.Errorf("failed to create the temporary directories: %v", err)
		return
	}

	// TODO: Check if directories were actually created

	if !wb.tempDirsCreated {
		t.Error("unexpected value for Workbook.tempDirsCreated, expected: true, found: false")
	}

	err = os.RemoveAll(wb.tempRootDir)
	if err != nil {
		t.Errorf("failed to remove the temporary directories: %v", err)
		return
	}
}
