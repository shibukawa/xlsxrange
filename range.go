package xlsxrange

import (
	"bytes"
	"fmt"
	"github.com/tealeg/xlsx"
	"strconv"
)

// Range struct treats range of spreadsheet cells
type Range struct {
	File       *xlsx.File  // Target file
	Sheet      *xlsx.Sheet // Target sheet
	Row        int         // Row number (1 origin)
	Column     int         // Column number (1 origin)
	NumRows    int         // Number of rows. AllRows means all rows.
	NumColumns int         // Number of cols. AllColumns means all columns.
}

const (
	AllRows    = -1 // All rows in sheet
	AllColumns = -1 // All columns in sheet
)

// NewWithFile creates Range instance
// It can accept notation parameters. See Range.Select for detail.
func NewWithFile(file *xlsx.File, notation ...interface{}) *Range {
	result := Range{
		File:       file,
		Sheet:      nil,
		Row:        1,
		Column:     1,
		NumRows:    AllRows,
		NumColumns: AllColumns,
	}
	if len(notation) > 0 {
		result.Select(notation...)
	}
	return &result
}

// New creates Range instance
// It can accept notation parameters. See Range.Select for detail.
func New(sheet *xlsx.Sheet, notation ...interface{}) *Range {
	result := Range{
		File:       sheet.File,
		Sheet:      sheet,
		Row:        1,
		Column:     1,
		NumRows:    AllRows,
		NumColumns: AllColumns,
	}
	if len(notation) > 0 {
		result.Select(notation...)
	}
	return &result
}

// SetSheet sets current sheet by name
func (r *Range) SetSheet(name string) error {
	sheet, ok := r.File.Sheet[name]
	if ok {
		r.Sheet = sheet
		return nil
	}
	return fmt.Errorf("Sheet name '%s' is missing", name)
}

// Select sets range.
// It allows the following styles:
//
// 	* row, col, numRows, numCols int (e.g. 10, 20, 3, 5)
// 	* row, col int    (e.g. 10, 20)
// 	* notation string (e.g. A2:B3)
func (r *Range) Select(notation ...interface{}) error {
	switch len(notation) {
	case 1:
		str, ok := notation[0].(string)
		if ok {
			sheetName, ranges, err := ParseA1Notation(str)
			if err != nil {
				return err
			}
			if sheetName != "" {
				sheet, ok := r.File.Sheet[sheetName]
				if !ok {
					return fmt.Errorf("Specified sheet is not found: %s", sheetName)
				}
				r.Sheet = sheet
			}
			r.Row = ranges[0]
			r.Column = ranges[1]
			r.NumRows = ranges[2]
			r.NumColumns = ranges[3]
		} else {
			return fmt.Errorf("Arguments should be string.")
		}
	case 2:
		row, ok1 := notation[0].(int)
		column, ok2 := notation[1].(int)
		if ok1 && ok2 {
			r.Row = row
			r.Column = column
			r.NumRows = 1
			r.NumColumns = 1
		} else {
			return fmt.Errorf("Arguments (row, column) should be integer.")
		}
	case 4:
		row, ok1 := notation[0].(int)
		column, ok2 := notation[1].(int)
		numRows, ok3 := notation[2].(int)
		numColumns, ok4 := notation[3].(int)
		if ok1 && ok2 && ok3 && ok4 {
			r.Row = row
			r.Column = column
			r.NumRows = numRows
			r.NumColumns = numColumns
		} else {
			return fmt.Errorf("Arguments (row, column, numRow, numColumns) should be integer.")

		}
	default:
		return fmt.Errorf("Select can accept single string or (int, int) or (int, int, int, int), but there are %d arguments", len(notation))
	}
	return nil
}

// Reset() clears selection
func (r *Range) Reset() {
	r.Row = 1
	r.Column = 1
	r.NumRows = AllRows
	r.NumColumns = AllColumns
}

// GetCell returns left top corner cell from selected range
func (r *Range) GetCell() *xlsx.Cell {
	return r.GetCellAt(0, 0)
}

// GetCellAt returns cell at relative location from selected range
//
// Input row, col are 0 origin. If selected position is D4 and input is 1, 1,
// this method returns cell at E5.
func (r *Range) GetCellAt(refRow, refCol int) *xlsx.Cell {
	return r.Sheet.Rows[r.Row+refRow-1].Cells[r.Column+refCol-1]
}

// GetCells returns cells in selected range
func (r *Range) GetCells() [][]*xlsx.Cell {
	rowCount := r.NumRows
	if rowCount == AllRows {
		rowCount = r.Sheet.MaxRow - r.Row + 1
	}
	if rowCount < 0 {
		rowCount = 0
	}
	rows := make([][]*xlsx.Cell, rowCount)
	columnCount := r.NumColumns
	if columnCount == AllColumns {
		columnCount = r.Sheet.MaxCol - r.Column + 1
	}
	for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
		row := make([]*xlsx.Cell, columnCount)
		rows[rowIndex] = row
		srcRow := r.Sheet.Rows[rowIndex+r.Row-1]

		for column := 0; column < columnCount; column++ {
			absCol := column + r.Column - 1
			row[column] = srcRow.Cells[absCol]
		}
	}
	return rows
}

func (r *Range) String() string {
	columnLabel := NumberToColumnStr(r.Column)
	if r.Row == -1 {
		return columnLabel
	}
	return fmt.Sprintf("%s%d:%s, %s", columnLabel, r.Row, columnLabel, columnLabel)
}

func (r *Range) Format(includeSheetName bool) string {
	var buffer bytes.Buffer
	if includeSheetName {
		buffer.WriteString(r.Sheet.Name)
		buffer.WriteByte('!')
	}
	if r.NumRows == AllRows && r.NumColumns == AllColumns {
		buffer.WriteString("1:1048576")
	} else if r.NumRows == AllRows {
		buffer.WriteString(NumberToColumnStr(r.Column))
		buffer.WriteByte(':')
		buffer.WriteString(NumberToColumnStr(r.Column + r.NumColumns - 1))
	} else if r.NumColumns == AllColumns {
		buffer.WriteString(strconv.Itoa(r.Row))
		buffer.WriteByte(':')
		buffer.WriteString(strconv.Itoa(r.Row + r.NumRows - 1))
	} else if r.NumColumns == 1 && r.NumRows == 1 {
		buffer.WriteString(NumberToColumnStr(r.Column))
		buffer.WriteString(strconv.Itoa(r.Row))
	} else {
		buffer.WriteString(NumberToColumnStr(r.Column))
		buffer.WriteString(strconv.Itoa(r.Row))
		buffer.WriteByte(':')
		buffer.WriteString(NumberToColumnStr(r.Column + r.NumColumns - 1))
		buffer.WriteString(strconv.Itoa(r.Row + r.NumRows - 1))
	}

	return buffer.String()
}
