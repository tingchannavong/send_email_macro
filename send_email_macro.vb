'Macro to send email based on certain default subject, body, recipients. With function to find last Friday date as needed by Weekly Report Standard

Function GetLatestFriday() As String
    Dim currentDate As Date
    Dim daysSinceLastFriday As Integer

    ' Get the current date
    currentDate = Date

    ' Calculate the number of days since the last Friday (considering Saturday as the first day)
    daysSinceLastFriday = Weekday(currentDate, vbSaturday)

    ' Determine the date of the last Friday
    Dim lastFriday As Date
    lastFriday = currentDate - daysSinceLastFriday

    ' Format the date as a string (adjust the format as needed)
    GetLatestFriday = Format(lastFriday, "yyyy-mm-dd")
End Function




Sub SendEmailWithLatestFriday()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim subjectLine As String

    ' Get the latest Friday in Subject Line
    subjectLine = "IT and Software Department Weekly Report - " & GetLatestFriday()

    ' Create multiline body text
    bodyText = "Hello Everyone," & vbNewLine & vbNewLine & _
                "Please find the attached document for our weekly report as of " & GetLatestFriday() & ". Let me know if there's any correction needed." & vbNewLine & vbNewLine & _
                "Yours Sincerely," & vbNewLine & _
                "YOUR NAME"

    ' Create Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")

    ' Create a new mail item
    Set OutlookMail = OutlookApp.CreateItem(0)

    ' Set email properties
    With OutlookMail
        .To = "RECIPIENT1@EMAIL.COM; RECIPIENT2@EMAIL.COM"
        .Subject = subjectLine
        .Body = bodyText

        ' Add attachments or additional settings as needed
        .Attachments.Add "YOUR_ATTACHMENT_FILE_PATH.pptx"

        ' Display the email (optional)
        .Display
        ' Send the email (comment out the Display line if you want to send it directly)
        '.Send
    End With

    ' Clean up
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

Private Sub SendReportMail_Click()

End Sub
