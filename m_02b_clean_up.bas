Attribute VB_Name = "m_02b_clean_up"
Option Explicit

'creator: Julian Jung (julian.jung@tws-partners.com)
'change-log:
    '2022-12-17: v01   - creation
    '2023-03-06: v01.1 - added no gridlines


Public Sub Clean_up()

    Dim int_zoom As Integer
    Dim wks_active As Worksheet
    Dim wks As Worksheet
    Dim rng_used As Range, rng As Range
    
    int_zoom = Application.InputBox("Zoom (in %)", Type:=1)
    Set wks_active = ActiveSheet
    
    For Each wks In Worksheets
        wks.Activate
        ActiveWindow.Zoom = int_zoom
        ActiveWindow.DisplayGridlines = False
        
        Range("A1").Select
        Set rng_used = wks.UsedRange
        For Each rng In rng_used
            If Not rng.Locked Then
                rng.Select
                Exit For
            End If
        Next rng
        
    Next wks
    
    wks_active.Activate

End Sub
