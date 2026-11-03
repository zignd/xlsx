package xlsx

// Cell represents a cell in a worksheet. A row has a collection of cells.
type Cell struct {
	row        *Row
	index      int
	identifier string
	// Key is the column key used to create the cell, if any.
	Key string
	// Value is the cell value. It is written when the owning row is committed.
	Value interface{}
}

// Identifier returns the A1-style reference of the cell (for example "B3").
func (c *Cell) Identifier() string {
	return c.identifier
}

// Index returns the zero-based column index of the cell.
func (c *Cell) Index() int {
	return c.index
}

// Row returns the row that owns the cell.
func (c *Cell) Row() *Row {
	return c.row
}
