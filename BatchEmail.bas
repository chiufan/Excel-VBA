Sub SendEmails()

    Dim olApp As Outlook.Application
    Dim olMail As MailItem
    Dim i As Integer
    Dim myEmailAttach As String

    Set olApp = New Outlook.Application

    i = 1
    myEmailAddress = ThisWorkbook.Sheets(1).Range("B1").Offset(i, 0)
    myEmailSubject = ThisWorkbook.Sheets(1).Range("C1").Offset(i, 0)
    myEmailBody = ThisWorkbook.Sheets(1).Range("D1").Offset(i, 0)
    myEmailAttach = ThisWorkbook.Sheets(1).Range("E1").Offset(i, 0)

    Do While Application.WorksheetFunction.IsText(myEmailAddress)
        Set olMail = olApp.CreateItem(olMailItem)

        With olMail
            .To = myEmailAddress
            .Subject = myEmailSubject
            .Body = myEmailBody
            If myEmailAttach = "" Then
            Else
                .Attachments.Add myEmailAttach
            End If
            If ThisWorkbook.Sheets(1).Range("I1") = "Manual" Then
                .Display
            Else
                .Send
            End If
        End With

        Set olMail = Nothing
        i = i + 1
        myEmailAddress = ThisWorkbook.Sheets(1).Range("B1").Offset(i, 0)
        myEmailSubject = ThisWorkbook.Sheets(1).Range("C1").Offset(i, 0)
        myEmailBody = ThisWorkbook.Sheets(1).Range("D1").Offset(i, 0)
        myEmailAttach = ThisWorkbook.Sheets(1).Range("E1").Offset(i, 0)
    Loop

    Set olApp = Nothing

End Sub

Sub Log()
    i = 1
    DatabaseLastRow = Application.CountA(ThisWorkbook.Sheets(1).Range("A:A"))
    For a = 1 To DatabaseLastRow - 1
        lastrow = Application.CountA(ThisWorkbook.Sheets(2).Range("A:A"))
        ThisWorkbook.Sheets(2).Cells(lastrow + 1, 1) = Date
        ThisWorkbook.Sheets(2).Cells(lastrow + 1, 2) = ThisWorkbook.Sheets(1).Cells(i + 1, 1)
        ThisWorkbook.Sheets(2).Cells(lastrow + 1, 3) = ThisWorkbook.Sheets(1).Cells(i + 1, 2)
        ThisWorkbook.Sheets(2).Cells(lastrow + 1, 4) = ThisWorkbook.Sheets(1).Cells(i + 1, 3)
        ThisWorkbook.Sheets(2).Cells(lastrow + 1, 5) = ThisWorkbook.Sheets(1).Cells(i + 1, 4)
        ThisWorkbook.Sheets(2).Cells(lastrow + 1, 6) = ThisWorkbook.Sheets(1).Cells(i + 1, 5)
        i = i + 1
    Next a
    ThisWorkbook.Sheets(1).Range("A2:E65536").ClearContents
End Sub
