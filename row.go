package xlsx

import (
	"errors"
	"fmt"
	"time"
)

// Row represents a row in a worksheet. A worksheet has a collection of rows.
type Row struct {
	worksheet *Worksheet
	index     int
	cells     []*Cell
	cellsMap  map[string]*Cell
	committed bool
	// Height is the optional row height in points. Zero means the default.
	Height float64
}

// CellOptions has options used when creating a new cell.
type CellOptions struct {
	// Key is the optional column key.
	Key *string
	// Value is the initial cell value.
	Value interface{}
}

// Index returns the zero-based index of the row.
func (r *Row) Index() int {
	return r.index
}

// SetHeight sets the row height in points. Excel accepts 0 (default) through
// 409.
func (r *Row) SetHeight(height float64) error {
	if height < 0 || height > 409 {
		return fmt.Errorf("%w: height must be between 0 and 409", ErrInvalidRow)
	}
	r.Height = height
	return nil
}

// Append adds a series of values to the row. On keyed worksheets the values
// are mapped to columns in definition order.
func (r *Row) Append(values ...interface{}) error {
	if r.worksheet != nil && r.worksheet.columns != nil && len(values) > len(r.worksheet.columns) {
		return fmt.Errorf("%w: %d values for %d columns", ErrColumnsDefined, len(values), len(r.worksheet.columns))
	}
	for i, value := range values {
		var (
			cell *Cell
			err  error
		)
		if r.worksheet != nil && r.worksheet.columns != nil {
			cell, err = r.AddCellWithKey(r.worksheet.columns[i].Key)
		} else {
			cell, err = r.AddCell()
		}
		if err != nil {
			return err
		}
		cell.Value = value
	}
	return nil
}

func (r *Row) appendTyped(value interface{}) (*Cell, error) {
	cell, err := r.AddCell()
	if err != nil {
		return nil, err
	}
	cell.Value = value
	return cell, nil
}

// WriteString appends a string cell. It can only be used on worksheets without
// keyed columns; use AddCellWithKey for keyed worksheets.
func (r *Row) WriteString(value string) (*Cell, error) {
	return r.appendTyped(value)
}

// WriteInt appends an integer cell.
func (r *Row) WriteInt(value int64) (*Cell, error) {
	return r.appendTyped(value)
}

// WriteFloat appends a floating-point cell.
func (r *Row) WriteFloat(value float64) (*Cell, error) {
	return r.appendTyped(value)
}

// WriteBool appends a boolean cell.
func (r *Row) WriteBool(value bool) (*Cell, error) {
	return r.appendTyped(value)
}

// WriteTime appends a date/time cell.
func (r *Row) WriteTime(value time.Time) (*Cell, error) {
	return r.appendTyped(value)
}

// AddCell adds a new cell to the row. It can only be used on worksheets where
// columns were not defined.
func (r *Row) AddCell() (*Cell, error) {
	if r.committed {
		return nil, errors.New("can't add cells to a committed row")
	}

	if r.worksheet == nil {
		return nil, errors.New("row is not attached to a worksheet")
	}

	if r.worksheet.columns != nil {
		return nil, fmt.Errorf("%w: use AddCellWithKey instead", ErrColumnsDefined)
	}

	cellIndex := len(r.cells)
	if cellIndex >= r.worksheet.maxColumns() {
		return nil, fmt.Errorf("a worksheet supports at most %d columns", r.worksheet.maxColumns())
	}
	cell := &Cell{
		row:        r,
		index:      cellIndex,
		identifier: createIdentifierFromCoords(cellIndex, r.index),
	}

	r.cells = append(r.cells, cell)

	return cell, nil
}

// AddCellWithKey is just like AddCell but adds a key to the cell. It should be
// used when columns were defined on the worksheet.
func (r *Row) AddCellWithKey(key string) (*Cell, error) {
	if r.committed {
		return nil, errors.New("can't add cells to a committed row")
	}

	if r.worksheet == nil {
		return nil, errors.New("row is not attached to a worksheet")
	}

	if r.worksheet.columns == nil {
		return nil, fmt.Errorf("%w: call DefineColumns first", ErrNoColumns)
	}

	var cellIndex = -1
	for i := 0; i < len(r.worksheet.columns); i++ {
		if r.worksheet.columns[i].Key == key {
			cellIndex = i
			break
		}
	}
	if cellIndex == -1 {
		return nil, fmt.Errorf("undefined column named %s", key)
	}

	if _, exists := r.cellsMap[key]; exists {
		return nil, fmt.Errorf("cell for column %s was already added", key)
	}

	cell := &Cell{
		row:        r,
		index:      cellIndex,
		identifier: createIdentifierFromCoords(cellIndex, r.index),
		Key:        key,
	}

	r.cellsMap[key] = cell

	return cell, nil
}
