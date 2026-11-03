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
)
