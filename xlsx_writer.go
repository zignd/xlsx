package xlsx

// NewWorkbook creates a new workbook, which is the base for every XLSX file.
// The file is written when Commit is called.
func NewWorkbook(filePath string) *Workbook {
	return &Workbook{
		FilePath: filePath,
		tempDirs: &workbookTempDirs{},
	}
}
