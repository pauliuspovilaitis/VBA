Sub Send_Email(SetSht As String, ThisRow As Long)

    Dim myMail As Object
    Dim EmailSubject As String
    Dim EmailFrom As String
    Dim EmailTo As String
    Dim EmailCC As String
    Dim EmailBCC As String
    Dim EmailBody As String
    Dim Attachments As String


        On Error Resume Next
    
        EmailFrom = ThisWorkbook.Worksheets(SetSht).Range("B" & ThisRow).Value
        EmailTo = ThisWorkbook.Worksheets(SetSht).Range("C" & ThisRow).Value
        EmailCC = ThisWorkbook.Worksheets(SetSht).Range("D" & ThisRow).Value
        EmailBCC = ThisWorkbook.Worksheets(SetSht).Range("E" & ThisRow).Value
        EmailSubject = ThisWorkbook.Worksheets(SetSht).Range("F" & ThisRow).Value
        EmailBody = ThisWorkbook.Worksheets(SetSht).Range("G" & ThisRow).Value
        Attachments = ThisWorkbook.Worksheets(SetSht).Range("H" & ThisRow).Value
    
        Set myMail = CreateObject("CDO.Message")
        myMail.From = EmailFrom
        myMail.To = EmailTo
        myMail.CC = EmailCC
        myMail.BCC = EmailBCC
        myMail.Subject = EmailSubject
        myMail.TextBody = EmailBody
        myMail.AddAttachment Attachments
    
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "XXX.XX.XX.XX"
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        myMail.Configuration.Fields.Update
        myMail.Send
    
        Set myMail = Nothing

        If (Err.Number = 0) Then
            ThisWorkbook.Worksheets(SetSht).Range("A" & ThisRow).Value = "OK"
            MsgBox ("Send! all OK")
        Else
            ThisWorkbook.Worksheets(SetSht).Range("A" & ThisRow).Value = "FAILED - " & Err.Number & " " & Err.Description
            MsgBox ("We have a problem....!")
        End If
    
        On Error GoTo 0

End Sub
