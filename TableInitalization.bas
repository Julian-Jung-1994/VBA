Attribute VB_Name = "TableInitalization"
Option Explicit

Public intFirstRow As Integer
Public intLastRow As Integer
Public intFirstColumn As Integer
Public intLastColumn As Integer
Public rngFirstCell As Range
Public rngLastCell As Range
Public rngTableRange As Range


Sub TableInitialization()

' Inputs 'TODO Userform/ Inputbox
Const intRow As Integer = 3 'row that will be used to deterime i) the firstRow and ii) lastColumn
Const intColumn As Integer = 1 ' column that will be used to determine i) the firstColumn and ii) lastRow
'----------------------------------------------------------------------------------------------------------'

' first cell
intFirstRow = intRow
intFirstColumn = intColumn
Set rngFirstCell = Cells(intFirstRow, intFirstColumn)

' last cell
intLastRow = Cells(Rows.Count, intFirstColumn).End(xlUp).row
intLastColumn = Cells(intFirstRow, Columns.Count).End(xlToLeft).column
Set rngLastCell = Cells(intLastRow, intLastColumn)

' entire range
Set rngTableRange = Range(rngFirstCell.Address & ":" & rngLastCell.Address)

End Sub
