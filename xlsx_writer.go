package xlsx

// NewWorkbook creates a new workbook, which is the base for every XLSX file.
// The file is written when Commit is called.
func NewWorkbook(filePath string) *Workbook {
	workbook, err := NewWorkbookWithOptions(filePath, Options{})
	if err != nil {
		// Options{} is always valid; this panic only protects future changes.
		panic(err)
	}
	return workbook
}
