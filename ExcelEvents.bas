Attribute VB_Name = "ExcelEvents"
Option Explicit


Sub ExcelEventsOff()
' turn off excel "events" that slow down the code

    With Application
        .ScreenUpdating = False ' no screen updating while the makro is running (line by line execution viewable)
        .DisplayAlerts = False ' excel pop-ups surpressed (e.g., do you really want to overwrite file?)
        .DisplayStatusBar = False ' status bar in the lower left corner does not update anymore
        .Calculation = xlCalculationManual ' stops excels recalculation of entire workbook when a single cell is udpated
    End With

End Sub


Sub ExcelEventsOn()
' turn the excel "events" back on for the usual excel experience
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = True
        .Calculation = xlCalculationAutomatic
        .CutCopyMode = False
    End With

End Sub
