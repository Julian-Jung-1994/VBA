Attribute VB_Name = "m_email"
Option Explicit

'creator: Julian Jung (julian.jung@tws-partners.com)
'change-log:
    '2022-09-05: review
    '2022-12-17: completely new; with outlook library & html body & signature
    '2022-12-19: loop for e-mail excel template

'future versions:
    'see TODO
    'inclusion of variable parameters


Public Sub Email()
' makro for sending emails - only helpful for standardized emails (e.g., without changing text)
' maybe for list of recipients in excel sheet collected who should receive same e-mail
    
    Dim email_app As Outlook.Application
    Dim email_item As Outlook.MailItem
    
    Dim i As Integer
    Dim arr_email() As Variant
    
    arr_email = Range("D9:G" & Range("D9").End(xlDown).Row).Value2 'TODO robustness for more than one e-mail address 'TODO 9 as constant
    
    For i = LBound(arr_email, 1) To UBound(arr_email, 1)
            
        Set email_app = New Outlook.Application
        Set email_item = email_app.CreateItem(olMailItem)
        
        With email_item
            .Display
            .To = arr_email(i, 1)
            '.CC = ""
            '.BCC = ""
            .Subject = "Test"
            If arr_email(i, 4) <> VBA.vbNullString Then
                .Attachments.Add arr_email(i, 4) 'full path
            End If

            .HTMLBody = "<HTML><BODY><p>" & _
                        "Dear " & arr_email(i, 2) & " " & arr_email(i, 3) & "," & "<br> <br>" & VBA.vbNewLine & _
                        "This is my first VBA email from Excel." & "<br> <br>" & VBA.vbNewLine & _
                        "With best regards," & "<br>" & VBA.vbNewLine & _
                        "Julian" & _
                        "</p></BODY></HTML>" & _
            .HTMLBody 'needed for signature
    
            .Send
        End With
    Next i
            
End Sub


Private Sub Email_text()

    Const str_rng_start As String = "D26"
    Const str_new_line As String = "<br>"
    
    Dim rng_start As Range, rng_end As Range
    Dim rng As Range
    Dim str_email As String
    
    Set rng_start = Range(str_rng_start)
    Set rng_end = Range("D" & Rows.Count).End(xlUp)

    str_email = ""
    For Each rng In Range(rng_start, rng_end)
        str_email = str_email & str_new_line
    Next rng
    
End Sub
