xlsxrange
================

This package is helper package for github.com/tealeg/xlsx. This package provides the feature to handle range notations.

.. code-block:: go

   file, _ := xlsx.OpenFile("test.xlsx")

   aRange := xlsxrange.New(file.Sheet["Sheet1"])
   aRange.Select("E5:F6")

   cell := aRange.GetCell(1, 2) // Relative position from E5
   cells := aRange.GetCells()   // [][]*xlsx.Cell

License
-----------

MIT
