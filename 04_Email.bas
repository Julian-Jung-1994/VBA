Attribute VB_Name = "Email"
Option Explicit

Sub Email()

    Dim objOutlook As Object
    Dim objMail As Object

    ' e-mail initialization
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(olMailItem)
    
    ' create and send the e-mail
    With objMail
        .To = ""
        .CC = ""
        .Sobject = ""
        .Body = ""
        .Attachements.Add
        .Send
    End With

End Sub
