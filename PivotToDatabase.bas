Attribute VB_Name = "PivotToDatabase"
Option Explicit

Dim arrData() As Variant

Dim lngRows As Long
Dim intColumns As Long
Dim lngRecords As Long
Dim intAttributes As Integer

Dim i As Long
Dim j As Long


Sub PivotToArray()
    
    ' usual initialization
    Worksheets(1).Select
    Call TableInitialization
    
    ' database initialization
    lngRows = intLastRow - intFirstRow
    intColumns = intLastColumn - intFirstColumn
    lngRecords = lngRows * intColumns
    intAttributes = intColumns
    ReDim arrData(1 To lngRecords, 1 To intAttributes)
    
    ' filling up the array
    For j = 1 To intColumns
        For i = 1 To lngRows
            arrData(i + lngRows * (j - 1), 1) = Cells(intFirstRow + i, intFirstColumn)
            arrData(i + lngRows * (j - 1), 2) = Cells(intFirstRow, intFirstColumn + j)
            arrData(i + lngRows * (j - 1), 3) = Cells(intFirstRow + i, intFirstColumn + j)
        Next i
    Next j

End Sub


Sub ArrayToDatabase()

    Call PivotToArray
    
    ThisWorkbook.Worksheets.Add After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = "PivotToDatabase"
    
    For i = 1 To lngRecords
        For j = 1 To intAttributes
            Cells(i, j) = arrData(i, j)
        Next j
    Next i

End Sub

