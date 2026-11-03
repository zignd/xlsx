package xlsx

import (
	"bufio"
	"database/sql/driver"
	"errors"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"time"
)

const worksheetBufferSize = 64 * 1024

const (
	maxWorksheetColumns = 16_384
	maxWorksheetRows    = 1_048_576
)

// Worksheet represents a worksheet in a workbook.
type Worksheet struct {
	workbook          *Workbook
	id                int
	name              string
	fileName          string
	filePath          string
	committed         bool
	rowsCommittedOnce bool
	pendingRows       []*Row
	rowsCount         int
	started           bool
	columns           []*WorksheetColumn
	file              *os.File
	writer            *bufio.Writer
	err               error
}

// WorksheetColumn represents a column in a worksheet.
type WorksheetColumn struct {
	// Key identifies the column when calling AddCellWithKey.
	Key string
	// Value is the header text written in the first row.
	Value string
	// Width is the optional column width in Excel units. A value of zero
	// leaves the column at its default width.
	Width float64
}

// ID returns the one-based identifier of the worksheet.
func (ws *Worksheet) ID() int {
	return ws.id
}

// Name returns the worksheet name.
func (ws *Worksheet) Name() string {
	return ws.name
}

// FileName returns the name of the worksheet part inside the XLSX package.
func (ws *Worksheet) FileName() string {
	return ws.fileName
}

// DefineColumns defines the worksheet columns and writes the header row. It is
// optional; a worksheet can be used without columns by calling AddCell.
func (ws *Worksheet) DefineColumns(columns []*WorksheetColumn) error {
	if ws.err != nil {
		return ws.err
	}
	if ws.committed {
		return fmt.Errorf("%w: can't define columns", ErrWorksheetCommitted)
	}
	if ws.rowsCommittedOnce {
		return errors.New("can't define columns after rows have been committed")
	}
	if len(ws.pendingRows) > 0 {
		return errors.New("can't define columns while there are pending rows")
	}
	if len(columns) == 0 {
		return fmt.Errorf("%w: at least one column is required", ErrInvalidColumn)
	}
	if len(columns) > maxWorksheetColumns {
		return fmt.Errorf("%w: a worksheet supports at most %d columns", ErrInvalidColumn, maxWorksheetColumns)
	}

	seen := make(map[string]struct{}, len(columns))
	for i, column := range columns {
		if column == nil {
			return fmt.Errorf("%w: column %d is nil", ErrInvalidColumn, i)
		}
		if column.Key == "" {
			return fmt.Errorf("%w: column %d has an empty key", ErrInvalidColumn, i)
		}
		if _, exists := seen[column.Key]; exists {
			return fmt.Errorf("%w: duplicate key %q", ErrInvalidColumn, column.Key)
		}
		if column.Width < 0 {
			return fmt.Errorf("%w: negative width for column %q", ErrInvalidColumn, column.Key)
		}
		if column.Width > 255 {
			return fmt.Errorf("%w: width for column %q exceeds 255", ErrInvalidColumn, column.Key)
		}
		seen[column.Key] = struct{}{}
	}

	ws.columns = columns

	row, err := ws.AddRow()
	if err != nil {
		return fmt.Errorf("failed to create the header row: %w", err)
	}
	for _, column := range ws.columns {
		cell, err := row.AddCellWithKey(column.Key)
		if err != nil {
			return fmt.Errorf("failed to add the header cell for column %q: %w", column.Key, err)
		}
		cell.Value = column.Value
	}

	if err := ws.CommitRows(); err != nil {
		return fmt.Errorf("failed to commit the header row: %w", err)
	}

	return nil
}

// AddRow adds a new row to the worksheet. Rows are buffered in memory until
// CommitRows is called.
func (ws *Worksheet) AddRow() (*Row, error) {
	if ws.err != nil {
		return nil, ws.err
	}
	if ws.committed {
		return nil, fmt.Errorf("%w: can't add rows", ErrWorksheetCommitted)
	}
	if ws.rowsCount >= maxWorksheetRows {
		return nil, fmt.Errorf("a worksheet supports at most %d rows", maxWorksheetRows)
	}

	row := &Row{
		worksheet: ws,
		index:     ws.rowsCount,
	}
	ws.pendingRows = append(ws.pendingRows, row)
	ws.rowsCount++

	if ws.columns != nil {
		row.cellsMap = make(map[string]*Cell, len(ws.columns))
	}

	return row, nil
}

func (ws *Worksheet) ensureStarted() error {
	if ws.started {
		return nil
	}
	if ws.err != nil {
		return ws.err
	}
	if err := ws.workbook.ensureTempDirs(); err != nil {
		return err
	}
	return ws.start()
}

func (ws *Worksheet) start() error {
	if ws.started {
		return errors.New("can't start a worksheet more than once")
	}
	if !ws.workbook.tempDirsCreated {
		return errors.New("can't start a worksheet if the temporary directories were not created yet")
	}

	ws.filePath = filepath.Join(ws.workbook.tempDirs.xlWorksheets, ws.fileName)
	f, err := os.Create(ws.filePath)
	if err != nil {
		return fmt.Errorf("failed to create %s: %w", ws.filePath, err)
	}

	writer := bufio.NewWriterSize(f, worksheetBufferSize)
	cleanup := func(startErr error) error {
		_ = writer.Flush()
		_ = f.Close()
		_ = os.Remove(ws.filePath)
		return startErr
	}

	if _, err := writer.WriteString(worksheetStart); err != nil {
		return cleanup(fmt.Errorf("failed to write the worksheet header to %s: %w", ws.filePath, err))
	}

	if hasColumnWidths(ws.columns) {
		if _, err := writer.WriteString("<cols>"); err != nil {
			return cleanup(fmt.Errorf("failed to write the columns to %s: %w", ws.filePath, err))
		}
		for i, column := range ws.columns {
			if column.Width <= 0 {
				continue
			}
			if _, err := writer.WriteString(columnXML(i+1, column.Width)); err != nil {
				return cleanup(fmt.Errorf("failed to write a column to %s: %w", ws.filePath, err))
			}
		}
		if _, err := writer.WriteString("</cols>"); err != nil {
			return cleanup(fmt.Errorf("failed to finish the columns in %s: %w", ws.filePath, err))
		}
	}

	if _, err := writer.WriteString(worksheetDataStart); err != nil {
		return cleanup(fmt.Errorf("failed to write the worksheet data header to %s: %w", ws.filePath, err))
	}

	ws.file = f
	ws.writer = writer
	ws.started = true
	return nil
}

func hasColumnWidths(columns []*WorksheetColumn) bool {
	for _, column := range columns {
		if column.Width > 0 {
			return true
		}
	}
	return false
}

func (ws *Worksheet) writeCell(cell *Cell) error {
	value := cell.Value
	if value == nil {
		_, err := ws.writer.WriteString(blankCellXML(cell.identifier))
		return err
	}

	if valuer, ok := value.(driver.Valuer); ok {
		if isNilValue(valuer) {
			_, err := ws.writer.WriteString(blankCellXML(cell.identifier))
			return err
		}
		driverValue, err := valuer.Value()
		if err != nil {
			return fmt.Errorf("failed to retrieve the value of a driver.Valuer: %w", err)
		}
		value = driverValue
	}

	switch v := value.(type) {
	case nil:
		_, err := ws.writer.WriteString(blankCellXML(cell.identifier))
		return err
	case time.Time:
		_, err := ws.writer.WriteString(dateCellXML(cell.identifier, v))
		return err
	case string:
		_, err := ws.writer.WriteString(textCellXML(cell.identifier, v))
		return err
	case []byte:
		_, err := ws.writer.WriteString(textCellXML(cell.identifier, string(v)))
		return err
	case bool:
		_, err := ws.writer.WriteString(boolCellXML(cell.identifier, v))
		return err
	}

	rv := reflect.ValueOf(value)
	switch rv.Kind() {
	case reflect.String:
		_, err := ws.writer.WriteString(textCellXML(cell.identifier, rv.String()))
		return err
	case reflect.Bool:
		_, err := ws.writer.WriteString(boolCellXML(cell.identifier, rv.Bool()))
		return err
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		_, err := ws.writer.WriteString(numberCellXML(cell.identifier, strconv.FormatInt(rv.Int(), 10)))
		return err
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		_, err := ws.writer.WriteString(numberCellXML(cell.identifier, strconv.FormatUint(rv.Uint(), 10)))
		return err
	case reflect.Float32, reflect.Float64:
		f := rv.Float()
		if math.IsNaN(f) || math.IsInf(f, 0) {
			return fmt.Errorf("%w: %v is not a valid spreadsheet number", ErrUnsupportedType, f)
		}
		formatted := formatFloat(f)
		if rv.Kind() == reflect.Float32 {
			formatted = formatFloat32(float32(f))
		}
		_, err := ws.writer.WriteString(numberCellXML(cell.identifier, formatted))
		return err
	case reflect.Slice:
		if rv.Type().Elem().Kind() == reflect.Uint8 {
			_, err := ws.writer.WriteString(textCellXML(cell.identifier, string(rv.Bytes())))
			return err
		}
		return fmt.Errorf("%w: %s", ErrUnsupportedType, rv.Type().String())
	case reflect.Struct:
		return fmt.Errorf("%w: %s", ErrUnsupportedType, rv.Type().String())
	default:
		return fmt.Errorf("%w: %s", ErrUnsupportedType, rv.Type().String())
	}
}

func isNilValue(value interface{}) bool {
	rv := reflect.ValueOf(value)
	switch rv.Kind() {
	case reflect.Chan, reflect.Func, reflect.Interface, reflect.Map, reflect.Pointer, reflect.Slice:
		return rv.IsNil()
	default:
		return false
	}
}

func (ws *Worksheet) writeRow(row *Row) error {
	if _, err := ws.writer.WriteString(startRowXML(row.index)); err != nil {
		return err
	}

	if ws.columns == nil {
		for _, cell := range row.cells {
			if err := ws.writeCell(cell); err != nil {
				return err
			}
		}
	} else {
		for i, column := range ws.columns {
			cell, ok := row.cellsMap[column.Key]
			if !ok {
				if _, err := ws.writer.WriteString(blankCellXML(createIdentifierFromCoords(i, row.index))); err != nil {
					return err
				}
				continue
			}
			if err := ws.writeCell(cell); err != nil {
				return err
			}
		}
	}

	if _, err := ws.writer.WriteString(endRowXML()); err != nil {
		return err
	}
	row.committed = true
	return nil
}

// CommitRows writes the rows stored in memory to disk. It can be called
// repeatedly to stream data without keeping the whole worksheet in memory.
func (ws *Worksheet) CommitRows() error {
	if ws.err != nil {
		return ws.err
	}
	if ws.committed {
		return fmt.Errorf("%w: can't commit rows", ErrWorksheetCommitted)
	}
	if len(ws.pendingRows) == 0 {
		return ErrNoRows
	}
	if err := ws.ensureStarted(); err != nil {
		return fmt.Errorf("failed to start the worksheet: %w", err)
	}
	if ws.writer == nil {
		return errors.New("the worksheet writer is not available")
	}

	for _, row := range ws.pendingRows {
		if err := ws.writeRow(row); err != nil {
			return fmt.Errorf("failed to write row %d: %w", row.index+1, err)
		}
	}

	if err := ws.writer.Flush(); err != nil {
		return fmt.Errorf("failed to flush the worksheet: %w", err)
	}

	ws.pendingRows = nil
	ws.rowsCommittedOnce = true
	return nil
}

// Commit commits the worksheet, writing its footer and closing the underlying
// file. Workbook.Commit finalizes worksheets automatically; calling Commit on
// an already committed worksheet returns ErrWorksheetCommitted.
func (ws *Worksheet) Commit() error {
	if ws.err != nil {
		return ws.err
	}
	if ws.committed {
		return fmt.Errorf("%w: can't commit it twice", ErrWorksheetCommitted)
	}
	if len(ws.pendingRows) > 0 {
		return fmt.Errorf("%w: call CommitRows first", ErrPendingRows)
	}
	if err := ws.ensureStarted(); err != nil {
		return fmt.Errorf("failed to start the worksheet: %w", err)
	}
	if ws.writer == nil || ws.file == nil {
		return errors.New("the worksheet writer is not available")
	}

	if _, err := ws.writer.WriteString(worksheetDataEnd); err != nil {
		return fmt.Errorf("failed to write the worksheet footer: %w", err)
	}
	if err := ws.writer.Flush(); err != nil {
		return fmt.Errorf("failed to flush the worksheet: %w", err)
	}
	if err := ws.file.Close(); err != nil {
		return fmt.Errorf("failed to close %s: %w", ws.filePath, err)
	}

	ws.file = nil
	ws.writer = nil
	ws.committed = true
	return nil
}

func (ws *Worksheet) closeFile() {
	if ws.writer != nil {
		_ = ws.writer.Flush()
	}
	if ws.file != nil {
		_ = ws.file.Close()
	}
	ws.writer = nil
	ws.file = nil
}

func validateSheetName(name string) error {
	if name == "" {
		return fmt.Errorf("%w: name cannot be empty", ErrInvalidSheetName)
	}
	if len([]rune(name)) > 31 {
		return fmt.Errorf("%w: name %q exceeds 31 characters", ErrInvalidSheetName, name)
	}
	for _, r := range name {
		if r < 0x20 || r == 0x7f {
			return fmt.Errorf("%w: name %q contains a control character", ErrInvalidSheetName, name)
		}
	}
	if strings.ContainsAny(name, `[]:*?/\`) {
		return fmt.Errorf("%w: name %q contains an invalid character", ErrInvalidSheetName, name)
	}
	if strings.HasPrefix(name, "'") || strings.HasSuffix(name, "'") {
		return fmt.Errorf("%w: name %q cannot start or end with an apostrophe", ErrInvalidSheetName, name)
	}
	if strings.EqualFold(name, "History") {
		return fmt.Errorf("%w: %q is reserved by Excel", ErrInvalidSheetName, name)
	}
	return nil
}
