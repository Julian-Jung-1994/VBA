Attribute VB_Name = "m_99_excel_events"
Option Explicit

'creator: Julian Jung (julian.jung@tws-partners.com)
'change-log:
    '2022-09-05: review


Private Sub Excel_events_off()
' turn off "excel events" that slow down or interfere

    With Application
        .ScreenUpdating = False 'no screen updating while the makro is running (line by line execution viewable)
        .Calculation = xlCalculationManual 'stops excels recalculation of entire workbook when a single cell is udpated
        .DisplayAlerts = False 'excel pop-ups surpressed (e.g., do you really want to overwrite file?)
        .EnableEvents = False 'excel events via other macros
        .DisplayStatusBar = False 'status bar in the lower left corner does not update anymore
    End With

End Sub


Private Sub Excel_events_on()
' turn the excel "events" back on for the usual excel experience

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .EnableEvents = True
        .DisplayStatusBar = True
        .CutCopyMode = False
    End With

End Sub
