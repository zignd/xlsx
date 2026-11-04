package xlsx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

// CellType describes the type of a cell returned by the reader.
type CellType uint8

const (
	CellTypeBlank CellType = iota
	CellTypeString
	CellTypeNumber
	CellTypeBool
	CellTypeDate
	CellTypeError
)

// CellData is a cell returned by a SheetReader.
type CellData struct {
	Ref       string
	Column    int
	Row       int
	Type      CellType
	Raw       string
	Number    float64
	BoolValue bool
	TimeValue time.Time
	Formula   string
	StyleID   int
	HasValue  bool
}

// String returns the most natural string representation of the cell.
func (c CellData) String() (string, error) {
	switch c.Type {
	case CellTypeString, CellTypeNumber, CellTypeError:
		return c.Raw, nil
	case CellTypeBool:
		if c.BoolValue {
			return "true", nil
		}
		return "false", nil
	case CellTypeDate:
		return c.TimeValue.Format(time.RFC3339), nil
	case CellTypeBlank:
		return "", nil
	default:
		return "", fmt.Errorf("%w: %d", ErrWrongCellType, c.Type)
	}
}

// Float64 returns the cell as a float64.
func (c CellData) Float64() (float64, error) {
	switch c.Type {
	case CellTypeNumber, CellTypeDate:
		value, err := strconv.ParseFloat(c.Raw, 64)
		if err != nil {
			return 0, fmt.Errorf("%w: %q is not a number", ErrWrongCellType, c.Raw)
		}
		return value, nil
	case CellTypeBool:
		if c.BoolValue {
			return 1, nil
		}
		return 0, nil
	default:
		return 0, fmt.Errorf("%w: %d", ErrWrongCellType, c.Type)
	}
}

// Int64 returns the cell as an int64.
func (c CellData) Int64() (int64, error) {
	value, err := c.Float64()
	if err != nil {
		return 0, err
	}
	if value != float64(int64(value)) {
		return 0, fmt.Errorf("%w: %v is not an integer", ErrWrongCellType, value)
	}
	return int64(value), nil
}

// Bool returns the cell as a boolean.
func (c CellData) Bool() (bool, error) {
	switch c.Type {
	case CellTypeBool:
		return c.BoolValue, nil
	case CellTypeNumber:
		switch c.Raw {
		case "1":
			return true, nil
		case "0":
			return false, nil
		}
	case CellTypeString:
		switch strings.ToLower(strings.TrimSpace(c.Raw)) {
		case "true", "1":
			return true, nil
		case "false", "0":
			return false, nil
		}
	}
	return false, fmt.Errorf("%w: %d", ErrWrongCellType, c.Type)
}

// Time returns the cell as a date/time.
func (c CellData) Time() (time.Time, error) {
	if c.Type != CellTypeDate {
		return time.Time{}, fmt.Errorf("%w: %d", ErrWrongCellType, c.Type)
	}
	return c.TimeValue, nil
}

// RowData is one row returned by a SheetReader.
type RowData struct {
	Index int
	Cells []CellData
}

// Cell returns the cell with the given A1 reference.
func (r *RowData) Cell(ref string) (*CellData, bool) {
	ref = strings.ToUpper(strings.TrimSpace(strings.ReplaceAll(ref, "$", "")))
	for i := range r.Cells {
		if r.Cells[i].Ref == ref {
			return &r.Cells[i], true
		}
	}
	return nil, false
}

// CellAt returns the cell at the zero-based column index.
func (r *RowData) CellAt(column int) (*CellData, bool) {
	for i := range r.Cells {
		if r.Cells[i].Column == column {
			return &r.Cells[i], true
		}
	}
	return nil, false
}

type xmlRow struct {
	R     int       `xml:"r,attr"`
	Cells []xmlCell `xml:"c"`
}

type xmlCell struct {
	R  string `xml:"r,attr"`
	T  string `xml:"t,attr"`
	S  int    `xml:"s,attr"`
	V  string `xml:"v"`
	F  string `xml:"f"`
	IS struct {
		T string `xml:"t"`
		R []struct {
			T string `xml:"t"`
		} `xml:"r"`
	} `xml:"is"`
}

func (c xmlCell) inlineText() string {
	var b strings.Builder
	b.WriteString(c.IS.T)
	for _, run := range c.IS.R {
		b.WriteString(run.T)
	}
	return b.String()
}

// SheetReader streams rows from a worksheet.
type SheetReader struct {
	reader  *Reader
	info    SheetInfo
	rc      io.ReadCloser
	dec     *xml.Decoder
	lastRow int
	current *RowData
	err     error
	closed  bool
}

// Info returns the worksheet metadata.
func (s *SheetReader) Info() SheetInfo {
	return s.info
}

// Next advances to the next row. It returns false at the end of the sheet or
// after an error; check Err in that case.
func (s *SheetReader) Next() bool {
	if s.closed || s.err != nil {
		return false
	}
	if err := s.reader.checkContext(); err != nil {
		s.err = err
		return false
	}

	for {
		token, err := s.dec.Token()
		if errors.Is(err, io.EOF) {
			return false
		}
		if err != nil {
			s.err = fmt.Errorf("%w: worksheet %q: %w", ErrCorruptFile, s.info.Name, err)
			return false
		}
		switch element := token.(type) {
		case xml.StartElement:
			if element.Name.Local != "row" {
				continue
			}
			var row xmlRow
			if err := s.dec.DecodeElement(&row, &element); err != nil {
				s.err = fmt.Errorf("%w: worksheet %q: %w", ErrCorruptFile, s.info.Name, err)
				return false
			}
			parsed, err := s.buildRow(row)
			if err != nil {
				s.err = err
				return false
			}
			s.current = parsed
			return true
		case xml.EndElement:
			if element.Name.Local == "sheetData" {
				return false
			}
		}
	}
}

// Row returns the most recently read row. The returned value is only valid
// until the next call to Next.
func (s *SheetReader) Row() *RowData {
	return s.current
}

// Err returns the first read error.
func (s *SheetReader) Err() error {
	return s.err
}

// Close releases the worksheet stream.
func (s *SheetReader) Close() error {
	if s.closed {
		return nil
	}
	s.closed = true
	if s.rc != nil {
		return s.rc.Close()
	}
	return nil
}

func (s *SheetReader) buildRow(row xmlRow) (*RowData, error) {
	rowIndex := row.R
	if rowIndex == 0 {
		rowIndex = s.lastRow + 1
		for _, cell := range row.Cells {
			if cell.R != "" {
				_, parsedRow, err := parseCellReference(cell.R)
				if err == nil && parsedRow > 0 {
					rowIndex = parsedRow
				}
				break
			}
		}
	}
	if rowIndex < 1 {
		return nil, fmt.Errorf("%w: worksheet %q has an invalid row number", ErrCorruptFile, s.info.Name)
	}
	if rowIndex > s.reader.options.MaxRows {
		return nil, fmt.Errorf("%w: worksheet %q exceeds %d rows", ErrReadLimit, s.info.Name, s.reader.options.MaxRows)
	}
	s.lastRow = rowIndex

	parsed := &RowData{Index: rowIndex, Cells: make([]CellData, 0, len(row.Cells))}
	columnCursor := -1
	for _, rawCell := range row.Cells {
		column := columnCursor + 1
		if rawCell.R != "" {
			parsedColumn, _, err := parseCellReference(rawCell.R)
			if err != nil {
				return nil, fmt.Errorf("%w: worksheet %q: %w", ErrCorruptFile, s.info.Name, err)
			}
			column = parsedColumn
		}
		if column >= s.reader.options.MaxColumns {
			return nil, fmt.Errorf("%w: worksheet %q exceeds %d columns", ErrReadLimit, s.info.Name, s.reader.options.MaxColumns)
		}

		cell, err := s.buildCell(rawCell, column, rowIndex)
		if err != nil {
			return nil, err
		}
		parsed.Cells = append(parsed.Cells, cell)
		columnCursor = column
	}
	return parsed, nil
}

func (s *SheetReader) buildCell(raw xmlCell, column, row int) (CellData, error) {
	ref := raw.R
	if ref == "" {
		ref = createIdentifierFromCoords(column, row-1)
	} else {
		ref = strings.ToUpper(strings.ReplaceAll(ref, "$", ""))
	}

	cell := CellData{
		Ref:     ref,
		Column:  column,
		Row:     row,
		StyleID: raw.S,
		Formula: strings.TrimSpace(raw.F),
		Raw:     raw.V,
	}

	switch raw.T {
	case "s":
		index, err := strconv.Atoi(strings.TrimSpace(raw.V))
		if err != nil {
			return CellData{}, fmt.Errorf("%w: invalid shared string index %q", ErrCorruptFile, raw.V)
		}
		value, err := s.reader.shared.lookup(index)
		if err != nil {
			return CellData{}, err
		}
		if err := s.checkCellLength(value); err != nil {
			return CellData{}, err
		}
		cell.Type = CellTypeString
		cell.Raw = value
		cell.HasValue = true
	case "inlineStr":
		value := raw.inlineText()
		if err := s.checkCellLength(value); err != nil {
			return CellData{}, err
		}
		cell.Type = CellTypeString
		cell.Raw = value
		cell.HasValue = value != "" || raw.IS.T != "" || len(raw.IS.R) > 0
	case "str":
		if err := s.checkCellLength(raw.V); err != nil {
			return CellData{}, err
		}
		cell.Type = CellTypeString
		cell.HasValue = true
	case "b":
		cell.Type = CellTypeBool
		cell.BoolValue = raw.V == "1" || strings.EqualFold(raw.V, "true")
		cell.HasValue = raw.V != ""
	case "e":
		cell.Type = CellTypeError
		cell.HasValue = raw.V != ""
	case "d":
		value, err := time.Parse(time.RFC3339, strings.TrimSpace(raw.V))
		if err != nil {
			return CellData{}, fmt.Errorf("%w: invalid date %q", ErrCorruptFile, raw.V)
		}
		cell.Type = CellTypeDate
		cell.TimeValue = value
		cell.HasValue = true
	case "n", "":
		if strings.TrimSpace(raw.V) == "" {
			cell.Type = CellTypeBlank
			return cell, nil
		}
		number, err := strconv.ParseFloat(strings.TrimSpace(raw.V), 64)
		if err != nil {
			return CellData{}, fmt.Errorf("%w: invalid number %q", ErrCorruptFile, raw.V)
		}
		cell.Number = number
		cell.HasValue = true
		if s.reader.styles.isDate(raw.S) {
			cell.Type = CellTypeDate
			cell.TimeValue = excelSerialToTime(number, s.reader.date1904)
		} else {
			cell.Type = CellTypeNumber
		}
	default:
		if err := s.checkCellLength(raw.V); err != nil {
			return CellData{}, err
		}
		cell.Type = CellTypeString
		cell.HasValue = raw.V != ""
	}

	return cell, nil
}

func (s *SheetReader) checkCellLength(value string) error {
	if utf8.RuneCountInString(value) > s.reader.options.MaxCellLength {
		return fmt.Errorf("%w: worksheet %q cell exceeds %d characters",
			ErrReadLimit, s.info.Name, s.reader.options.MaxCellLength)
	}
	return nil
}
