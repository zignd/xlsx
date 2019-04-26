package xlsx

// NewWorkbook creates a new workbook, which is the base for every XLSX file.
func NewWorkbook(filePath string) *Workbook {
	return &Workbook{
		FilePath: filePath,
		tempDirs: &workbookTempDirs{},
	}
}
