Attribute VB_Name = "D_Form"
Option Explicit


Sub Form()
    
    Dim arrEntries() As Variant
    Dim i As Integer
    Dim lngRow As Long
        
    Const attributes As Long = 8

    ' determine the length of the array
    ReDim arrEntries(1 To attributes) As Variant
    
    ' make sure to be on the right sheet
    Worksheets("Form").Activate
    
    ' fill the array with the entered data
    For i = 1 To 6
        arrEntries(i) = Range("B" & 2 * i + 1).Value
    Next i
    
    For i = 7 To 8
        arrEntries(i) = Range("E" & (i - 6) * 2 + 5)
    Next i
    
    ' switch to the data sheet
    Worksheets("Data").Activate

    ' determine the next entry row underneath the other entries
    lngRow = Range("A" & Rows.Count).End(xlUp).Offset(1).row
    
    ' deliver the data from the array to the database
    For i = 1 To attributes
        Cells(lngRow, i) = arrEntries(i)
    Next i
    
End Sub



