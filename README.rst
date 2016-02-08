xlsxrange
================

This package is helper package for github.com/tealeg/xlsx. This package provides the feature to handle range notations.

.. code-block:: go

   file, _ := xlsx.OpenFile("test.xlsx")

   range := xlsxrange.New(file.Sheet["Sheet1"])
   range.Select("E5:F6")

   cell := range.GetCell(1, 2) // Relative position from E5
   cells := range.GetCells()   // [][]*xlsx.Cell

License
-----------

MIT
