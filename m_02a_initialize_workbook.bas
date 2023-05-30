Attribute VB_Name = "m_02a_initialize_workbook"
Option Explicit

' author: Julian Jung
' change-log
    '2022-01-24: v01 - creation

Sub Initialize_wkb()
Attribute Initialize_wkb.VB_ProcData.VB_Invoke_Func = " \n14"

    Const str_theme_path As String = "" 'TODO path and file with theme
    Dim wks As Worksheet

    ' excel events off
    Application.Run ("Excel_events_off") 'TODO make sure to have macro in your workbook

    ' apply theme
    ActiveWorkbook.ApplyTheme (str_theme_path)

    ' gridlines & zoom
    With ActiveWindow
        .DisplayGridlines = False
        .Zoom = 70
    End With

    ' row heights & column widths
    Rows.RowHeight = 15 ' standard
    Rows(3).RowHeight = 30 ' for header 1
    Rows(6).RowHeight = 25 ' for header 1
    Columns.ColumnWidth = 12 ' standard
    Columns("A:C").ColumnWidth = 2 ' first two columns + header
    Columns("Q:R").ColumnWidth = 2 ' last two columns

    ' working window (hide rows and columns)
    Rows("51:" & Rows.Count).Hidden = True
    Range(Range("S1"), Cells(1, Columns.Count)).EntireColumn.Hidden = True

    ' header fonts - for both headers
    With Range("C3:P3, C6:P6")
        With .Font
        .Size = 14
        .Color = VBA.vbWhite
        .Bold = True
        End With
        
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' interior colors for both headers
    Range("C3:P3").Interior.ThemeColor = xlThemeColorAccent3
    Range("C6:P6").Interior.ThemeColor = xlThemeColorAccent5

    ' duplicate designed sheet to have 3 new sheets
    Sheets(1).Activate
    ActiveSheet.Copy After:=ActiveSheet
    ActiveSheet.Copy After:=ActiveSheet
    
    ' worksheet tab color
    With Worksheets(1)
        .Name = "Overview"
        .Tab.ThemeColor = xlThemeColorAccent3
    End With
    
    With Worksheets(2)
        .Name = "Analysis"
        .Tab.ThemeColor = xlThemeColorAccent5
    End With
    
    With Worksheets(3)
        .Name = "Data"
        .Tab.ThemeColor = xlThemeColorDark1
        .Tab.TintAndShade = -0.149998474074526
    End With
    
    ' A1 and first sheets
    For Each wks In Worksheets
        wks.Activate
        Range("C3").Select
    Next wks
    Worksheets(1).Activate
    
    ' excel events on
    Application.Run ("Excel_events_on") 'TODO make sure to have macro in your workbook

End Sub
