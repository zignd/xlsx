package xlsx

import (
	"archive/zip"
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
	"unicode/utf8"
)

const (
	maxWorksheetColumns = 16_384
	maxWorksheetRows    = 1_048_576
	maxRowHeight        = 409
	maxFormulaLength    = 8_192
	maxHyperlinkLength  = 2_079
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
	hyperlinks        []hyperlink
	mergeCells        []string
	freezeRows        int
	freezeCols        int
	autoFilter        string
	unflushedRows     int
}

type hyperlink struct {
	ref    string
	target string
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

func (ws *Worksheet) maxRows() int {
	return ws.workbook.options.MaxRows
}

func (ws *Worksheet) maxColumns() int {
	return ws.workbook.options.MaxColumns
}

func (ws *Worksheet) maxCellLength() int {
	return ws.workbook.options.MaxCellLength
}

// DefineColumns defines the worksheet columns and writes the header row. It is
// optional; a worksheet can be used without columns by calling AddCell.
func (ws *Worksheet) DefineColumns(columns []*WorksheetColumn) error {
	if err := ws.workbook.checkContext(); err != nil {
		return err
	}
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
	if len(columns) > ws.maxColumns() {
		return fmt.Errorf("%w: a worksheet supports at most %d columns", ErrInvalidColumn, ws.maxColumns())
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
	if err := ws.workbook.checkContext(); err != nil {
		return nil, err
	}
	if ws.err != nil {
		return nil, ws.err
	}
	if ws.committed {
		return nil, fmt.Errorf("%w: can't add rows", ErrWorksheetCommitted)
	}
	if ws.rowsCount >= ws.maxRows() {
		return nil, fmt.Errorf("a worksheet supports at most %d rows", ws.maxRows())
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

// MergeCells merges one or more A1 ranges. Ranges must be valid and must not
// overlap.
func (ws *Worksheet) MergeCells(ranges ...string) error {
	if ws.err != nil {
		return ws.err
	}
	if ws.committed {
		return fmt.Errorf("%w: can't merge cells", ErrWorksheetCommitted)
	}

	for _, ref := range ranges {
		normalized, rect, err := normalizeRange(ref, ws.maxRows(), ws.maxColumns())
		if err != nil {
			return err
		}
		for _, existing := range ws.mergeCells {
			_, existingRect, _ := normalizeRange(existing, ws.maxRows(), ws.maxColumns())
			if rangesOverlap(rect, existingRect) {
				return fmt.Errorf("%w: range %s overlaps %s", ErrInvalidRange, normalized, existing)
			}
		}
		ws.mergeCells = append(ws.mergeCells, normalized)
	}
	return nil
}

// FreezePanes freezes the given number of rows and columns. It must be called
// before any rows are committed.
func (ws *Worksheet) FreezePanes(rows, columns int) error {
	if ws.err != nil {
		return ws.err
	}
	if ws.committed {
		return fmt.Errorf("%w: can't freeze panes", ErrWorksheetCommitted)
	}
	if ws.started {
		return errors.New("can't change freeze panes after rows have been written")
	}
	if rows < 0 || columns < 0 || rows > ws.maxRows() || columns > ws.maxColumns() {
		return fmt.Errorf("%w: freeze pane counts are out of range", ErrInvalidRange)
	}
	if rows == 0 && columns == 0 {
		return fmt.Errorf("%w: at least one frozen row or column is required", ErrInvalidRange)
	}
	ws.freezeRows = rows
	ws.freezeCols = columns
	return nil
}

// SetAutoFilter sets the auto-filter range for the worksheet.
func (ws *Worksheet) SetAutoFilter(ref string) error {
	if ws.err != nil {
		return ws.err
	}
	if ws.committed {
		return fmt.Errorf("%w: can't set an auto-filter", ErrWorksheetCommitted)
	}
	normalized, _, err := normalizeRange(ref, ws.maxRows(), ws.maxColumns())
	if err != nil {
		return err
	}
	ws.autoFilter = normalized
	return nil
}

func (ws *Worksheet) ensureStarted() error {
	if ws.started {
		return nil
	}
	if ws.err != nil {
		return ws.err
	}
	if ws.workbook.streaming {
		return ws.startStreaming()
	}
	if err := ws.workbook.ensureTempDirs(); err != nil {
		return err
	}
	return ws.start()
}

func (ws *Worksheet) writePreamble(writer *bufio.Writer) error {
	if _, err := writer.WriteString(worksheetStartPrefix); err != nil {
		return err
	}
	if _, err := writer.WriteString(worksheetViewsXML(ws.freezeRows, ws.freezeCols)); err != nil {
		return err
	}
	if _, err := writer.WriteString(`<sheetFormatPr defaultRowHeight="15"/>`); err != nil {
		return err
	}
	if hasColumnWidths(ws.columns) {
		if _, err := writer.WriteString("<cols>"); err != nil {
			return err
		}
		for i, column := range ws.columns {
			if column.Width <= 0 {
				continue
			}
			if _, err := writer.WriteString(columnXML(i+1, column.Width)); err != nil {
				return err
			}
		}
		if _, err := writer.WriteString("</cols>"); err != nil {
			return err
		}
	}
	if _, err := writer.WriteString(worksheetDataStart); err != nil {
		return err
	}
	return nil
}

func (ws *Worksheet) startStreaming() error {
	if ws.started {
		return errors.New("can't start a worksheet more than once")
	}
	state := ws.workbook.stream
	if state == nil || state.closed {
		return errors.New("streaming writer is closed")
	}
	if state.openSheet != nil {
		return fmt.Errorf("%w: worksheet %q is still open", ErrStreamingOrder, state.openSheet.name)
	}

	header := &zip.FileHeader{Name: "xl/worksheets/" + ws.fileName, Method: zip.Deflate}
	if ws.workbook.options.Compression == CompressionStore {
		header.Method = zip.Store
	}
	header.SetMode(ws.workbook.options.FileMode)
	entry, err := state.zipWriter.CreateHeader(header)
	if err != nil {
		return fmt.Errorf("failed to create streaming worksheet %s: %w", ws.fileName, err)
	}

	writer := bufio.NewWriterSize(entry, ws.workbook.options.BufferSize)
	if err := ws.writePreamble(writer); err != nil {
		return fmt.Errorf("failed to write the streaming worksheet header: %w", err)
	}

	ws.writer = writer
	ws.file = nil
	ws.started = true
	ws.filePath = ""
	state.openSheet = ws
	return nil
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

	writer := bufio.NewWriterSize(f, ws.workbook.options.BufferSize)
	cleanup := func(startErr error) error {
		_ = writer.Flush()
		_ = f.Close()
		_ = os.Remove(ws.filePath)
		return startErr
	}

	if err := ws.writePreamble(writer); err != nil {
		return cleanup(fmt.Errorf("failed to write the worksheet header to %s: %w", ws.filePath, err))
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

func (ws *Worksheet) writeStringCell(identifier, value string, styleID int) error {
	if utf8.RuneCountInString(value) > ws.maxCellLength() {
		return fmt.Errorf("%w: maximum is %d characters", ErrCellTooLong, ws.maxCellLength())
	}
	if ws.workbook.options.SharedStrings {
		index := ws.workbook.sharedStrings.index(value)
		_, err := ws.writer.WriteString(sharedStringCellXML(identifier, index, styleID))
		return err
	}
	_, err := ws.writer.WriteString(textCellXML(identifier, value, styleID))
	return err
}

func (ws *Worksheet) writeCell(cell *Cell) error {
	if cell.Hyperlink != "" {
		target, err := normalizeHyperlink(cell.Hyperlink)
		if err != nil {
			return err
		}
		ws.hyperlinks = append(ws.hyperlinks, hyperlink{ref: cell.identifier, target: target})
	}

	styleID := cell.StyleID
	if styleID != 0 {
		if err := ws.workbook.validateStyleID(styleID); err != nil {
			return err
		}
	}

	if cell.Formula != "" {
		formula, err := normalizeFormula(cell.Formula)
		if err != nil {
			return err
		}
		_, err = ws.writer.WriteString(formulaCellXML(cell.identifier, formula, styleID))
		return err
	}

	value := cell.Value
	if value == nil {
		_, err := ws.writer.WriteString(blankCellXML(cell.identifier, styleID))
		return err
	}

	if valuer, ok := value.(driver.Valuer); ok {
		if isNilValue(valuer) {
			_, err := ws.writer.WriteString(blankCellXML(cell.identifier, styleID))
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
		_, err := ws.writer.WriteString(blankCellXML(cell.identifier, styleID))
		return err
	case time.Time:
		_, err := ws.writer.WriteString(dateCellXML(cell.identifier, v, styleID))
		return err
	case string:
		return ws.writeStringCell(cell.identifier, v, styleID)
	case []byte:
		return ws.writeStringCell(cell.identifier, string(v), styleID)
	case bool:
		_, err := ws.writer.WriteString(boolCellXML(cell.identifier, v, styleID))
		return err
	}

	rv := reflect.ValueOf(value)
	switch rv.Kind() {
	case reflect.String:
		return ws.writeStringCell(cell.identifier, rv.String(), styleID)
	case reflect.Bool:
		_, err := ws.writer.WriteString(boolCellXML(cell.identifier, rv.Bool(), styleID))
		return err
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		_, err := ws.writer.WriteString(numberCellXML(cell.identifier, strconv.FormatInt(rv.Int(), 10), styleID))
		return err
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		_, err := ws.writer.WriteString(numberCellXML(cell.identifier, strconv.FormatUint(rv.Uint(), 10), styleID))
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
		_, err := ws.writer.WriteString(numberCellXML(cell.identifier, formatted, styleID))
		return err
	case reflect.Slice:
		if rv.Type().Elem().Kind() == reflect.Uint8 {
			return ws.writeStringCell(cell.identifier, string(rv.Bytes()), styleID)
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
	if row.Height < 0 || row.Height > maxRowHeight {
		return fmt.Errorf("%w: height must be between 0 and %d", ErrInvalidRow, maxRowHeight)
	}
	if _, err := ws.writer.WriteString(startRowXML(row.index, row.Height)); err != nil {
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
				if _, err := ws.writer.WriteString(blankCellXML(createIdentifierFromCoords(i, row.index), 0)); err != nil {
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
	if err := ws.workbook.checkContext(); err != nil {
		return err
	}
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
		if err := ws.workbook.checkContext(); err != nil {
			return err
		}
		if err := ws.writeRow(row); err != nil {
			return fmt.Errorf("failed to write row %d: %w", row.index+1, err)
		}
	}

	if err := ws.writer.Flush(); err != nil {
		return fmt.Errorf("failed to flush the worksheet: %w", err)
	}

	committedRows := len(ws.pendingRows)
	ws.pendingRows = nil
	ws.rowsCommittedOnce = true

	if ws.workbook.streaming {
		ws.unflushedRows += committedRows
		interval := ws.workbook.options.FlushEveryRows
		if interval > 0 && ws.unflushedRows >= interval {
			if err := ws.workbook.Flush(); err != nil {
				return err
			}
			ws.unflushedRows = 0
		}
	}
	return nil
}

// Commit commits the worksheet, writing its footer and closing the underlying
// file. Workbook.Commit finalizes worksheets automatically; calling Commit on
// an already committed worksheet returns ErrWorksheetCommitted.
func (ws *Worksheet) Commit() error {
	if err := ws.workbook.checkContext(); err != nil {
		return err
	}
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
	if ws.writer == nil || (!ws.workbook.streaming && ws.file == nil) {
		return errors.New("the worksheet writer is not available")
	}

	if _, err := ws.writer.WriteString(worksheetDataEnd); err != nil {
		return fmt.Errorf("failed to write the worksheet footer: %w", err)
	}
	if _, err := ws.writer.WriteString(autoFilterXML(ws.autoFilter)); err != nil {
		return fmt.Errorf("failed to write the auto-filter: %w", err)
	}
	if _, err := ws.writer.WriteString(mergeCellsXML(ws.mergeCells)); err != nil {
		return fmt.Errorf("failed to write merged cells: %w", err)
	}
	if _, err := ws.writer.WriteString(hyperlinksXML(ws.hyperlinks)); err != nil {
		return fmt.Errorf("failed to write hyperlinks: %w", err)
	}
	if _, err := ws.writer.WriteString(worksheetEnd); err != nil {
		return fmt.Errorf("failed to write the worksheet footer: %w", err)
	}
	if err := ws.writer.Flush(); err != nil {
		return fmt.Errorf("failed to flush the worksheet: %w", err)
	}
	if ws.workbook.streaming {
		if ws.workbook.stream != nil && ws.workbook.stream.openSheet == ws {
			ws.workbook.stream.openSheet = nil
		}
	} else {
		if err := ws.file.Close(); err != nil {
			return fmt.Errorf("failed to close %s: %w", ws.filePath, err)
		}
	}

	if err := ws.writeWorksheetRelationships(); err != nil {
		return err
	}

	ws.file = nil
	ws.writer = nil
	ws.committed = true
	return nil
}

func (ws *Worksheet) writeWorksheetRelationships() error {
	if len(ws.hyperlinks) == 0 {
		return nil
	}
	var b strings.Builder
	b.WriteString(worksheetRelsStart)
	for i, link := range ws.hyperlinks {
		if strings.HasPrefix(link.target, "#") {
			b.WriteString(relationshipXML(i+1, contentTypeHyperlinkR, link.target))
		} else {
			b.WriteString(externalRelationshipXML(i+1, contentTypeHyperlinkR, link.target))
		}
	}
	b.WriteString(worksheetRelsEnd)

	filePath := filepath.Join(ws.workbook.tempDirs.xlWorksheetRels, ws.fileName+".rels")
	if ws.workbook.streaming {
		return ws.workbook.writeStreamPart("xl/worksheets/_rels/"+ws.fileName+".rels", b.String())
	}
	if err := os.WriteFile(filePath, []byte(b.String()), ws.workbook.options.FileMode); err != nil {
		return fmt.Errorf("failed to write worksheet relationships to %s: %w", filePath, err)
	}
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

func normalizeFormula(formula string) (string, error) {
	formula = strings.TrimPrefix(strings.TrimSpace(formula), "=")
	if formula == "" {
		return "", fmt.Errorf("%w: formula cannot be empty", ErrInvalidFormula)
	}
	if len([]rune(formula)) > maxFormulaLength {
		return "", fmt.Errorf("%w: formula exceeds %d characters", ErrInvalidFormula, maxFormulaLength)
	}
	for _, r := range formula {
		if r < 0x20 || r == 0x7f {
			return "", fmt.Errorf("%w: formula contains a control character", ErrInvalidFormula)
		}
	}
	return formula, nil
}

func normalizeHyperlink(target string) (string, error) {
	target = strings.TrimSpace(target)
	if target == "" {
		return "", fmt.Errorf("%w: target cannot be empty", ErrInvalidHyperlink)
	}
	if len([]rune(target)) > maxHyperlinkLength {
		return "", fmt.Errorf("%w: target exceeds %d characters", ErrInvalidHyperlink, maxHyperlinkLength)
	}
	for _, r := range target {
		if r < 0x20 || r == 0x7f {
			return "", fmt.Errorf("%w: target contains a control character", ErrInvalidHyperlink)
		}
	}
	return target, nil
}
