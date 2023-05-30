Attribute VB_Name = "m_03b_cell_selection"
Option Explicit

'creator: Julian Jung (julian.jung@tws-partners.com)
'change-log:
    '2022-12-17: v01 - creation


Public Sub Select_by_color()

    Dim rng_select As Range, rng As Range
    Dim rng_color As Range
    Dim lng_color As Long


    Set rng_color = Application.InputBox("Legend cell with color", Type:=8)
    lng_color = rng_color.DisplayFormat.Interior.Color
    
    For Each rng In ActiveSheet.UsedRange
        If rng.DisplayFormat.Interior.Color = lng_color And _
           rng.Address <> rng_color.Address Then
            
            If rng_select Is Nothing Then
                Set rng_select = rng
            Else
                Set rng_select = Application.Union(rng_select, rng)
            End If
            
        End If
    Next rng

    rng_select.Select

End Sub
