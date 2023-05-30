Attribute VB_Name = "pp_CopyDiagram"
Option Explicit

Sub Powerpoint_CopyDiagram()

' This module contains the rudimentary code to copy a diagram from an excel workbook to a powerpoint presentation.


    Dim pp As Powerpoint.Application
    Dim pPres As Powerpoint.Presentation
    Dim pSlide As Powerpoint.Slide
    Dim pShape As Powerpoint.Shape

    ' start powerpoint object
    Set pp = New Powerpoint.Application
    
    ' create pp presentation
    Set pPres = pp.Presentations.Add

    ' create new slide
    Set pSlide = pPres.Slides.AddSlide(pPres.Slides.Count + 1, _
                                       pPres.SlideMaster.CustomLayouts(7))  '7 gives empty slide
    
    ' copy diagram from excel
    ThisWorkbook.Worksheets("Diagram").ChartObjects("ppChart").Copy

    ' paste diagram to Powerpoint
    pSlide.Shapes.Paste
    
    ' more or less proper size of the diagram
    Set pShape = pSlide.Shapes(1)
    With pShape
        .LockAspectRatio = msoTrue
        .Left = 100
        .Top = 40
        .Width = 800
    End With

End Sub
