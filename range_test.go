package xlsxrange

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"testing"
)

func createFile() *xlsx.File {
	file := &xlsx.File{
		Sheet: make(map[string]*xlsx.Sheet),
	}
	for s := 1; s < 4; s++ {
		sheet := &xlsx.Sheet{
			File: file,
			Name: fmt.Sprintf("Sheet %d", s),
		}
		file.Sheets = append(file.Sheets, sheet)
		cols := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}
		for rowNum := 1; rowNum < 16; rowNum++ {
			row := &xlsx.Row{
				Sheet: sheet,
				Cells: make([]*xlsx.Cell, len(cols)),
			}
			for i, col := range cols {
				row.Cells[i] = &xlsx.Cell{}
				row.Cells[i].SetString(fmt.Sprintf("%s%d", col, rowNum))
			}
			sheet.Rows = append(sheet.Rows, row)
		}
		file.Sheet[sheet.Name] = sheet
		sheet.MaxRow = len(sheet.Rows)
		sheet.MaxCol = len(cols)
	}

	return file
}

func TestNewRange(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"])

	if aRange.Row != 1 {
		t.Errorf("uninitialized row should be 1, but %d\n", aRange.Row)
	}
	if aRange.Column != 1 {
		t.Errorf("uninitialized column should be 1, but %d\n", aRange.Column)
	}
	if aRange.NumRows != AllRows {
		t.Errorf("uninitialized number of rows should be all rows, but %d\n", aRange.NumRows)
	}
	if aRange.NumColumns != AllColumns {
		t.Errorf("uninitialized number of columns should be all columns, but %d\n", aRange.NumColumns)
	}
}

func TestSelectByR1C1(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], 5, 4, 3, 2)

	if aRange.Row != 5 {
		t.Errorf("row should be 5, but %d\n", aRange.Row)
	}
	if aRange.Column != 4 {
		t.Errorf("column should be 4, but %d\n", aRange.Column)
	}
	if aRange.NumRows != 3 {
		t.Errorf("number of rows should be 3, but %d\n", aRange.NumRows)
	}
	if aRange.NumColumns != 2 {
		t.Errorf("number of columns should be 2, but %d\n", aRange.NumColumns)
	}
}

func TestSelectSingleCellByR1C1(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], 5, 4)

	if aRange.Row != 5 {
		t.Errorf("row should be 5, but %d\n", aRange.Row)
	}
	if aRange.Column != 4 {
		t.Errorf("column should be 4, but %d\n", aRange.Column)
	}
	if aRange.NumRows != 1 {
		t.Errorf("number of rows should be 1, but %d\n", aRange.NumRows)
	}
	if aRange.NumColumns != 1 {
		t.Errorf("number of columns should be 1, but %d\n", aRange.NumColumns)
	}
}

func TestSelectByA1Notation(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], "D5:F6")

	if aRange.Row != 5 {
		t.Errorf("row should be 5, but %d\n", aRange.Row)
	}
	if aRange.Column != 4 {
		t.Errorf("column should be 4, but %d\n", aRange.Column)
	}
	if aRange.NumRows != 2 {
		t.Errorf("number of rows should be 2, but %d\n", aRange.NumRows)
	}
	if aRange.NumColumns != 3 {
		t.Errorf("number of columns should be 3, but %d\n", aRange.NumColumns)
	}
}

func TestSelectByA1NotationAllRow(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], "D:F")

	if aRange.Row != 1 {
		t.Errorf("row should be 1, but %d\n", aRange.Row)
	}
	if aRange.Column != 4 {
		t.Errorf("column should be 4, but %d\n", aRange.Column)
	}
	if aRange.NumRows != AllRows {
		t.Errorf("number of rows should be AllRows, but %d\n", aRange.NumRows)
	}
	if aRange.NumColumns != 3 {
		t.Errorf("number of columns should be 3, but %d\n", aRange.NumColumns)
	}
}

func TestSelectByA1NotationAllColumn(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], "5:6")

	if aRange.Row != 5 {
		t.Errorf("row should be 5, but %d\n", aRange.Row)
	}
	if aRange.Column != 1 {
		t.Errorf("column should be 1, but %d\n", aRange.Column)
	}
	if aRange.NumRows != 2 {
		t.Errorf("number of rows should be 2, but %d\n", aRange.NumRows)
	}
	if aRange.NumColumns != AllColumns {
		t.Errorf("number of columns should be AllColumns, but %d\n", aRange.NumColumns)
	}
}

func TestSelectByA1NotationWithSheetName(t *testing.T) {
	file := createFile()
	aRange := NewWithFile(file, "Sheet 2!D5:F6")

	if aRange.Sheet.Name != "Sheet 2" {
		t.Errorf("")
	}
	if aRange.Row != 5 {
		t.Errorf("row should be 5, but %d\n", aRange.Row)
	}
	if aRange.Column != 4 {
		t.Errorf("column should be 4, but %d\n", aRange.Column)
	}
	if aRange.NumRows != 2 {
		t.Errorf("number of rows should be 2, but %d\n", aRange.NumRows)
	}
	if aRange.NumColumns != 3 {
		t.Errorf("number of columns should be 3, but %d\n", aRange.NumColumns)
	}
}

func TestReset(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], 5, 4, 3, 2)

	aRange.Reset()
	if aRange.Row != 1 {
		t.Errorf("reseted row should be 1, but %d\n", aRange.Row)
	}
	if aRange.Column != 1 {
		t.Errorf("reseted column should be 1, but %d\n", aRange.Column)
	}
	if aRange.NumRows != AllRows {
		t.Errorf("reseted number of rows should be all rows, but %d\n", aRange.NumRows)
	}
	if aRange.NumColumns != AllColumns {
		t.Errorf("reseted number of columns should be all columns, but %d\n", aRange.NumColumns)
	}
}

func TestGetCell(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"])

	aRange.Row = 1
	aRange.Column = 1

	if aRange.GetCell(0, 0).Value != "A1" {
		t.Errorf("aRange.GetCell(0, 0) from A1 should be A1, but %s", aRange.GetCell(0, 0).Value)
	}
	if aRange.GetCell(1, 1).Value != "B2" {
		t.Errorf("aRange.GetCell(1, 1) from A1 should be B2, but %s", aRange.GetCell(0, 0).Value)
	}

	aRange.Row = 2
	aRange.Column = 2

	if aRange.GetCell(0, 0).Value != "B2" {
		t.Errorf("aRange.GetCell(0, 0) from A1 should be B2, but %s", aRange.GetCell(0, 0).Value)
	}
	if aRange.GetCell(1, 1).Value != "C3" {
		t.Errorf("aRange.GetCell(1, 1) from A1 should be C3, but %s", aRange.GetCell(0, 0).Value)
	}
}

func TestGetCells(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], 5, 4, 3, 2)

	cells := aRange.GetCells()

	if len(cells) != 3 {
		t.Errorf("row count should be 3, but %d", len(cells))
	}
	if len(cells[0]) != 2 {
		t.Errorf("col count should be 2, but %d", len(cells[0]))
	}
	if cells[0][0].Value != "D5" {
		t.Errorf("cells[0][0] should be 'D5', but %s", cells[0][0].Value)
	}
}

func TestFormatWithMultipleCells(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], 5, 4, 2, 3)
	if aRange.Format(false) != "D5:F6" {
		t.Errorf("Range.Format(false) should return 'D5:F6', but %s", aRange.Format(false))
	}
	if aRange.Format(true) != "Sheet 1!D5:F6" {
		t.Errorf("Range.Format(false) should return 'Sheet 1!D5:F6', but %s", aRange.Format(false))
	}
}

func TestFormatWithSingleCell(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], 5, 4)
	if aRange.Format(false) != "D5" {
		t.Errorf("Range.Format(false) should return 'D5', but %s", aRange.Format(false))
	}
	if aRange.Format(true) != "Sheet 1!D5" {
		t.Errorf("Range.Format(false) should return 'Sheet 1!D5', but %s", aRange.Format(false))
	}
}

func TestFormatWithAllRows(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], 1, 4, AllRows, 1)
	if aRange.Format(false) != "D:D" {
		t.Errorf("Range.Format(false) should return 'D:D', but %s", aRange.Format(false))
	}
	if aRange.Format(true) != "Sheet 1!D:D" {
		t.Errorf("Range.Format(false) should return 'Sheet 1!D:D', but %s", aRange.Format(false))
	}
}

func TestFormatWithAllColumns(t *testing.T) {
	file := createFile()
	aRange := New(file.Sheet["Sheet 1"], 5, 1, 1, AllColumns)
	if aRange.Format(false) != "5:5" {
		t.Errorf("Range.Format(false) should return '5:5', but %s", aRange.Format(false))
	}
	if aRange.Format(true) != "Sheet 1!5:5" {
		t.Errorf("Range.Format(false) should return 'Sheet 1!5:5', but %s", aRange.Format(false))
	}
}
