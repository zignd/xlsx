package xlsx

import (
	"errors"
	"fmt"
)

// Row represents a row in a worksheet. A worksheet has a collection of rows.
type Row struct {
	worksheet *Worksheet
	index     int
	cells     []*Cell
	cellsMap  map[string]*Cell
	committed bool
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
	if cellIndex >= maxWorksheetColumns {
		return nil, fmt.Errorf("a worksheet supports at most %d columns", maxWorksheetColumns)
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
