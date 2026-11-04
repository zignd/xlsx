package xlsx

import (
	"strings"
	"time"
)

// Cell represents a cell in a worksheet. A row has a collection of cells.
type Cell struct {
	row        *Row
	index      int
	identifier string
	// Key is the column key used to create the cell, if any.
	Key string
	// Value is the cell value. It is written when the owning row is committed.
	Value interface{}
	// Formula is an optional formula without the leading "=".
	Formula string
	// Hyperlink is an optional external or internal hyperlink target.
	Hyperlink string
	// StyleID is the cell style ID returned by Workbook.AddStyle. Zero means
	// the default style; dates automatically use the date style.
	StyleID int
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

// SetString sets a string value.
func (c *Cell) SetString(value string) *Cell {
	c.Value = value
	return c
}

// SetInt sets an integer value.
func (c *Cell) SetInt(value int64) *Cell {
	c.Value = value
	return c
}

// SetFloat sets a floating-point value.
func (c *Cell) SetFloat(value float64) *Cell {
	c.Value = value
	return c
}

// SetBool sets a boolean value.
func (c *Cell) SetBool(value bool) *Cell {
	c.Value = value
	return c
}

// SetTime sets a date/time value.
func (c *Cell) SetTime(value time.Time) *Cell {
	c.Value = value
	return c
}

// SetFormula sets a formula, removing an optional leading "=".
func (c *Cell) SetFormula(formula string) *Cell {
	c.Formula = strings.TrimPrefix(strings.TrimSpace(formula), "=")
	return c
}

// SetHyperlink sets an external URL or an internal reference such as
// "#Sheet2!A1".
func (c *Cell) SetHyperlink(target string) *Cell {
	c.Hyperlink = target
	return c
}

// SetStyle sets the cell style ID.
func (c *Cell) SetStyle(styleID int) *Cell {
	c.StyleID = styleID
	return c
}
