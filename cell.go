package xlsx

// Cell represents a cell in a worksheet. A row has a collection of cells.
type Cell struct {
	row        *Row
	index      int
	identifier string
	Key        string
	Value      interface{}
}
