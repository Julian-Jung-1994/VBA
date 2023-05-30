Attribute VB_Name = "m_01a_lock_cells"
Option Explicit

Public Sub Lock_Workbook_with_Legend_Cell()
    
    Dim wks As Worksheet
    Dim rng_input As Range, rng_all As Range, rng As Range
    Dim lng_color As Long
    Dim bool_cancel As Boolean
    
    bool_cancel = False

    For Each wks In Worksheets
        If bool_cancel Then Exit Sub
        
        wks.Activate
        bool_cancel = Lock_Worksheet_with_Legend_Cell
    Next wks
    
    Worksheets(1).Activate
    
End Sub


Public Function Lock_Worksheet_with_Legend_Cell() As Boolean
    
    Dim wks As Worksheet
    Dim rng_all As Range, rng_input As Range, rng As Range
    Dim lng_color As Long
    Dim bool_cancel As Boolean

    bool_cancel = False

    Set wks = ActiveSheet
    Set rng_all = wks.UsedRange
    ActiveSheet.Cells.Locked = True
                
    On Error GoTo input_cancelled
        Set rng_input = Application.InputBox("Cell with Color", Type:=8)
    
    On Error GoTo 0
    lng_color = rng_input.DisplayFormat.Interior.Color
    
    If lng_color <> 16777215 Then
        For Each rng In rng_all
            If rng.DisplayFormat.Interior.Color = lng_color Then
                rng.Select
                Selection.Locked = False
            End If
        Next
    End If
        
    rng_input.Locked = True
    Range("A1").Select
    
    Exit Function
    
input_cancelled:
    bool_cancel = True
    
    Lock_Worksheet_with_Legend_Cell = bool_cancel
    
End Function
