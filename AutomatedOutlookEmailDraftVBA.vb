Sub AutomatedOutlookEmailDraft()
    Call SendEmailWithAttachment("abc@xyz.com", "dxm@xyz.com; tvl@xyz.com; plv@xyz.com; qaq@xyz.com ")
End Sub

Sub SendEmailWithAttachment(mailTo As String, mailCC As String)
    Dim OutApp As Object
    Dim OutMail As Object
    Dim mailSubject As String
    Dim mailBody As String

    ' Create the Outlook application and email item
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    ' Define email parameters
    mailSubject = "B6KLM-19963: 8300016934-NML-"
    
    ' Compose the email body in HTML format
    mailBody = "<p style='font-family: Calibri; font-size: 11pt;'>" & _
               "<span style='color: #A6A6A6;'>Our ref.#: 2883-JDN-ALS-E-</span><br><br>" & _
               "Dear Maria,<br><br>" & _
               "Please find attached the summary of the XXX submission.<br><br>" & _
               "For the complete documentation, please refer to the link below:<br><br>" & _
               "Kindly submit the documents via Aconex.<br><br><br>" & _
               "</p>"

    ' Configure the email
    With OutMail
        .To = mailTo
        .CC = mailCC
        .Subject = mailSubject
        .HTMLBody = mailBody   ' Use HTMLBody to format the email in HTML
        ' Attach files manually before sending using Outlook interface
        .Display   ' Use .Send to send the email without previewing
    End With

    ' Clean up
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub




