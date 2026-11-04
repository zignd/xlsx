package xlsx

import "errors"

// Sentinel errors returned by the package. They can be inspected with
// [errors.Is].
var (
	// ErrWorkbookCommitted is returned when an operation is attempted on a
	// workbook that has already been committed.
	ErrWorkbookCommitted = errors.New("workbook already committed")

	// ErrWorksheetCommitted is returned when an operation is attempted on a
	// worksheet that has already been committed.
	ErrWorksheetCommitted = errors.New("worksheet already committed")

	// ErrNoWorksheets is returned when a workbook without worksheets is
	// committed.
	ErrNoWorksheets = errors.New("a workbook needs at least one worksheet")

	// ErrNoRows is returned when CommitRows is called without pending rows.
	ErrNoRows = errors.New("there are no rows to commit")

	// ErrPendingRows is returned when a worksheet is committed while it still
	// has rows waiting to be written.
	ErrPendingRows = errors.New("worksheet has pending rows")

	// ErrNoColumns is returned when a keyed cell is added to a worksheet that
	// has no columns defined.
	ErrNoColumns = errors.New("no columns were defined on the worksheet")

	// ErrColumnsDefined is returned when an unkeyed cell is added to a
	// worksheet that uses keyed columns.
	ErrColumnsDefined = errors.New("columns were defined on the worksheet")

	// ErrUnsupportedType is returned when a cell value cannot be represented
	// in an XLSX file.
	ErrUnsupportedType = errors.New("unsupported cell value type")

	// ErrInvalidSheetName is returned when a worksheet name is not a valid
	// Excel sheet name.
	ErrInvalidSheetName = errors.New("invalid worksheet name")

	// ErrDuplicateSheetName is returned when two worksheets share a name.
	ErrDuplicateSheetName = errors.New("duplicate worksheet name")

	// ErrInvalidColumn is returned when a column definition is invalid.
	ErrInvalidColumn = errors.New("invalid worksheet column")

	// ErrInvalidStyle is returned when a style or style ID is invalid.
	ErrInvalidStyle = errors.New("invalid cell style")

	// ErrInvalidFormula is returned when a formula is invalid.
	ErrInvalidFormula = errors.New("invalid cell formula")

	// ErrInvalidHyperlink is returned when a hyperlink is invalid.
	ErrInvalidHyperlink = errors.New("invalid cell hyperlink")

	// ErrInvalidRange is returned when an A1 range is invalid.
	ErrInvalidRange = errors.New("invalid A1 range")

	// ErrInvalidRow is returned when a row property is invalid.
	ErrInvalidRow = errors.New("invalid row")

	// ErrCellTooLong is returned when a text value exceeds the cell limit.
	ErrCellTooLong = errors.New("cell value is too long")

	// ErrTooManySheets is returned when a workbook exceeds its sheet limit.
	ErrTooManySheets = errors.New("too many worksheets")

	// ErrReadLimit is returned when a reader hits a configured safety limit.
	ErrReadLimit = errors.New("reader limit exceeded")

	// ErrCorruptFile is returned when an XLSX package is malformed.
	ErrCorruptFile = errors.New("corrupt XLSX file")

	// ErrEncrypted is returned for password-protected workbooks.
	ErrEncrypted = errors.New("encrypted workbooks are not supported")

	// ErrUnsupportedFormat is returned for formats such as XLSB.
	ErrUnsupportedFormat = errors.New("unsupported spreadsheet format")

	// ErrSheetNotFound is returned when a requested sheet does not exist.
	ErrSheetNotFound = errors.New("worksheet not found")

	// ErrSheetClosed is returned when reading from a closed worksheet.
	ErrSheetClosed = errors.New("worksheet reader is closed")

	// ErrInvalidCellReference is returned for malformed A1 references.
	ErrInvalidCellReference = errors.New("invalid cell reference")

	// ErrWrongCellType is returned when a typed cell accessor is used on the
	// wrong cell type.
	ErrWrongCellType = errors.New("wrong cell type")
)
