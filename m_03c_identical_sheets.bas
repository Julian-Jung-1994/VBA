Attribute VB_Name = "m_03c_identical_sheets"
Option Explicit

'creator: Julian Jung (julian.jung@tws-partners.com)
'change-log:
    '2022-12-16: v01 - creation
    

Public Sub Identical_Worksheets()
    
    Dim wks_original As Worksheet, wks_verify As Worksheet ', wks_log As Worksheet
    Dim rng_original As Range, rng_verify As Range
    Dim rng As Range
    Dim str_address As String
    
    Set wks_original = ActiveWorkbook.Worksheets(1)
    Set wks_verify = ActiveWorkbook.Worksheets(2)
    
'    'log sheet
'    Worksheets.Add After:=Worksheets(Worksheets.Count)
'    Set wks_log = Worksheets(Worksheets.Count)
    
    'loop through sheets and mark deviations with red color
    For Each rng In wks_original.UsedRange
        str_address = rng.Address
        Set rng_original = wks_original.Range(str_address)
        Set rng_verify = wks_verify.Range(str_address)
        
        If rng_original.Value2 <> rng_verify.Value2 Then
            rng_original.Interior.Color = VBA.vbRed
            rng_verify.Interior.Color = VBA.vbRed
        End If
    Next rng

    wks_original.Activate
    Range("A1").Select
    

End Sub
