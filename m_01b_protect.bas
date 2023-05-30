Attribute VB_Name = "m_01b_protect"
Option Explicit

'creator: Julian Jung (julian.jung@tws-partners.com)
'change-log:
    '2022-12-17: v01 - creation
    '2022-12-19: v02 - white cell selected, all cells selectable


Public Sub Protect_sheets_and_workbook()

    Dim str_password As String
    Dim wks As Worksheet
    
    str_password = Application.InputBox("Password", Type:=2)
    If StrPtr(str_password) = 0 Then Exit Sub
    
    For Each wks In ActiveWorkbook.Worksheets
        wks.Protect Password:=str_password
        wks.EnableSelection = xlNoRestrictions 'wks.EnableSelection = xlUnlockedCells
    Next wks
 
    ActiveWorkbook.Protect Password:=str_password
    
    Worksheets(1).Activate
    Range("A1").Select

End Sub


Public Sub Unprotect_sheets_and_workbook()

    Dim str_password As String
    Dim wks As Worksheet
    
    str_password = Application.InputBox("Password", Type:=2)
    If StrPtr(str_password) = 0 Then Exit Sub
    
    For Each wks In Worksheets
        wks.Unprotect Password:=str_password
    Next wks

    ActiveWorkbook.Unprotect Password:=str_password
    
    Worksheets(1).Activate
    Range("A1").Select

End Sub
