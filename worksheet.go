package xlsx

import (
	"database/sql/driver"
	"fmt"
	"os"
	"path"
	"reflect"
	"time"

	"github.com/davecgh/go-spew/spew"
	"github.com/pkg/errors"
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
}

// WorksheetColumn represents a column in a worksheet.
type WorksheetColumn struct {
	Key   string
	Value string
}

// DefineColumns defines the worksheet columns. It's optional.
func (ws *Worksheet) DefineColumns(columns []*WorksheetColumn) error {
	if ws.committed {
		return errors.New("can't define columns on a committed worksheet")
	}

	if ws.rowsCommittedOnce {
		return errors.New("can't define columns if rows have been committed")
	}

	ws.columns = columns

	row, err := ws.AddRow()
	if err != nil {
		return errors.Wrap(err, "failed to create the header row")
	}
	for i := 0; i < len(ws.columns); i++ {
		cell, err := row.AddCellWithKey(ws.columns[i].Key)
		cell.Value = ws.columns[i].Value
		if err != nil {
			return errors.Wrap(err, "failed to add one of the header's cells")
		}
	}

	err = ws.CommitRows()
	if err != nil {
		return errors.Wrap(err, "failed to commit the header row")
	}

	return nil
}

// AddRow adds a new row to the worksheet.
func (ws *Worksheet) AddRow() (*Row, error) {
	if ws.committed {
		return nil, errors.New("can't add rows to a committed worksheet")
	}

	row := &Row{
		worksheet: ws,
		index:     ws.rowsCount,
	}
	ws.pendingRows = append(ws.pendingRows, row)
	ws.rowsCount = ws.rowsCount + 1

	if ws.columns != nil {
		row.cellsMap = make(map[string]*Cell)
	}

	return row, nil
}

func (ws *Worksheet) start() error {
	if !ws.workbook.tempDirsCreated {
		return errors.New("can't start a worksheet if the temporary directories were not created yet")
	}

	if ws.started {
		return errors.New("can't start a worksheet more than once")
	}

	ws.filePath = path.Join(ws.workbook.tempDirs.xlWorksheets, ws.fileName)
	f, err := os.Create(ws.filePath)
	if err != nil {
		return errors.Wrapf(err, "failed to create %s", ws.filePath)
	}
	defer f.Close()

	_, err = f.WriteString(startWorksheet)
	if err != nil {
		return errors.Wrapf(err, "failed to append START_WORKSHEET to file %s", ws.filePath)
	}
	_, err = f.WriteString(startWorksheetData)
	if err != nil {
		return errors.Wrapf(err, "failed to append START_WORKSHEET_DATA to file %s", ws.filePath)
	}

	ws.started = true

	return nil
}

func (ws *Worksheet) createRow(row *Row) error {
	// TODO: Benchmark how error checking affects performance
	f, err := os.OpenFile(ws.filePath, os.O_APPEND|os.O_WRONLY, os.ModePerm)
	if err != nil {
		return errors.Wrapf(err, "failed to open file %s to append a new row", ws.filePath)
	}
	defer f.Close()

	_, err = f.WriteString(startRowFormat(row.index))
	if err != nil {
		return errors.Wrapf(err, "failed to append a new row to file %s", ws.filePath)
	}

	// TODO: Use reflection to check the type and create the appropriate kind of cell
	if ws.columns == nil {
		for i := 0; i < len(row.cells); i++ {
			cell := row.cells[i]
			f.WriteString(cellFormat(cell.identifier, fmt.Sprint(cell.Value)))
		}
	} else {
		lastCellIndex := 0
		for i := 0; i < len(ws.columns); i++ {
			cell, ok := row.cellsMap[ws.columns[i].Key]
			if ok {
				valuer, ok := cell.Value.(driver.Valuer)
				if ok {
					value, err := valuer.Value()
					if err != nil {
						return errors.Wrap(err, "failed to retrieve the Value of a driver.Valuer")
					}
					if value == nil {
						lastCellIndex++
						f.WriteString(cellFormat(createIdentifierFromCoords(lastCellIndex, row.index), ""))
						continue
					}
					cell.Value = value
				}

				t := reflect.TypeOf(cell.Value)
				k := t.Kind()
				switch k {
				case reflect.String, reflect.Bool:
					f.WriteString(cellFormat(cell.identifier, fmt.Sprint(cell.Value)))
				case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64, reflect.Float32, reflect.Float64:
					f.WriteString(numberCellFormat(cell.identifier, cell.Value))
				case reflect.Struct:
					if t.String() != "time.Time" {
						return errors.Errorf("%s is not supported in a cell", t.String())
					}
					f.WriteString(dateCellFormat(cell.identifier, cell.Value.(time.Time)))
				default:
					return errors.Errorf("%s is not supported in a cell", k.String())
				}

				lastCellIndex = cell.index
			} else {
				lastCellIndex++
				f.WriteString(cellFormat(createIdentifierFromCoords(lastCellIndex, row.index), ""))
			}
		}
	}

	_, err = f.WriteString(endRow)
	if err != nil {
		return errors.Wrapf(err, "failed to end row on file %s", ws.filePath)
	}

	return nil
}

func (ws *Worksheet) end() error {
	if !ws.started {
		return errors.New("can't end a worksheet if it has not been started yet")
	}

	f, err := os.OpenFile(ws.filePath, os.O_APPEND|os.O_WRONLY, os.ModePerm)
	if err != nil {
		return errors.Wrapf(err, "failed to open file %s to append a new row", ws.filePath)
	}
	defer f.Close()

	_, err = f.WriteString(endWorksheetData)
	if err != nil {
		return errors.Wrapf(err, "failed to append END_WORKSHEET_DATA to file %s", ws.filePath)
	}
	_, err = f.WriteString(endWorksheet)
	if err != nil {
		return errors.Wrapf(err, "failed to append END_WORKSHEET to file %s", ws.filePath)
	}

	return nil
}

// CommitRows commits rows stored in memory.
func (ws *Worksheet) CommitRows() error {
	if ws.committed {
		return errors.New("can't commit rows from a committed worksheet")
	}

	if len(ws.pendingRows) == 0 {
		return errors.New("there are no rows to commit")
	}

	if !ws.workbook.tempDirsCreated {
		err := ws.workbook.createTempDirs()
		if err != nil {
			return errors.Wrap(err, "failed to create temporary directories")
		}
	}

	if !ws.started {
		err := ws.start()
		if err != nil {
			return errors.Wrap(err, "failed to start the worksheet")
		}
	}

	for len(ws.pendingRows) > 0 {
		row := ws.pendingRows[0]
		ws.pendingRows = ws.pendingRows[1:]
		err := ws.createRow(row)
		if err != nil {
			return errors.Wrapf(err, "failed to create row %s", spew.Sprint(row))
		}
	}

	if !ws.rowsCommittedOnce {
		ws.rowsCommittedOnce = true
	}

	return nil
}

// Commit commits the worksheet.
func (ws *Worksheet) Commit() error {
	if ws.committed {
		return errors.New("can't commit an already committed worksheet")
	}

	if len(ws.pendingRows) > 0 {
		return errors.New("can't commit worksheet if there are still pending rows to be committed")
	}

	err := ws.end()
	if err != nil {
		return errors.Wrap(err, "failed to end the worksheet")
	}

	ws.committed = true

	return nil
}
