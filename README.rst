xlsxrange
================

This package is helper package for github.com/tealeg/xlsx. This package provides the feature to handle range notations.

.. code-block:: go

   file, _ := xlsx.OpenFile("test.xlsx")

   aRange := xlsxrange.New(file.Sheet["Sheet1"], "E5:F6")

   cell := aRange.GetCell(1, 2) // Relative position from E5
   cells := aRange.GetCells()   // [][]*xlsx.Cell

Usage
-----------

* ``xlsxrange.New(sheet *xlsx.Sheet, notation interface{}...)``
* ``xlsxrange.NewWithFile(file *xlsx.File, notation interface{}...)``

  These are constructor function to create ``Range`` instances.

  These functions can accept notation patterns as same as ``Range.Select()``

* ``Range.Select(notation interface{}...) error``

  Select range by parameters. It can accept three variations of notations:

  * R1C1 notation with row and column and row number and column number (four integers)
  * R1C1 notation with row and column (two integers)
  * A1 notation string

  .. code-block:: go

     aRange.Select(4, 5, 2, 2)
     aRange.Select(5, 6)
     aRange.Select("C4")

* ``Range.SetSheet(name string) error``

  Set current sheet by its name.

* ``Range.Reset()``

  Reset selection. 

* ``Range.GetCell(relRow, relCol int) *xlsx.Cell``

  It returns the cell by relative position from left top corner in selected range.
  (0, 0) is a left top corner.

* ``Range.GetCells() [][]*xlsx.Cell``

  It returns the all cells in selected range.

License
-----------

MIT
