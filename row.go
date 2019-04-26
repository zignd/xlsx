package xlsx

import (
	"github.com/pkg/errors"
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
	Key   *string
	Value interface{}
}

// AddCell adds a new cell to the row.
func (r *Row) AddCell() (*Cell, error) {
	if r.committed {
		return nil, errors.New("can't add cells to a committed row")
	}

	if r.worksheet.columns != nil {
		return nil, errors.New("can't add cells without keys to this worksheet as columns were defined")
	}

	cellIndex := len(r.cells)
	cell := &Cell{
		row:        r,
		index:      cellIndex,
		identifier: createIdentifierFromCoords(cellIndex, r.index),
	}

	r.cells = append(r.cells, cell)

	return cell, nil
}

// AddCellWithKey is just like AddCell but adds a key to the cell. Should be used if columns were defined on the worksheet.
func (r *Row) AddCellWithKey(key string) (*Cell, error) {
	if r.committed {
		return nil, errors.New("can't add cells to a committed row")
	}

	if r.worksheet.columns == nil {
		return nil, errors.New("can't add cells with keys if no columns were defined")
	}

	var cellIndex int
	var found bool
	for i := 0; i < len(r.worksheet.columns); i++ {
		if r.worksheet.columns[i].Key == key {
			cellIndex = i
			found = true
			break
		}
	}
	if !found {
		return nil, errors.Errorf("undefined column named %s", key)
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
