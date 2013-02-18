Sub FileReader()
    Application.ScreenUpdating = False

    ThisWorkbook.Sheets(1).Range("2:65536").Clear
    MyFolder = ThisWorkbook.Sheets(3).Range("B1")
    MyFileName = Dir(MyFolder & ThisWorkbook.Sheets(3).Range("B2"))

    Dim i, j As Integer
    Dim tempost() As String
    DataRange = ThisWorkbook.Sheets(3).Range("B3")
    ReDim tempost(1 To DataRange)

    For j = 1 To DataRange
        tempost(j) = ThisWorkbook.Sheets(2).Cells(j, 1)
    Next j
    i = 0
    Do Until MyFileName = ""
        With Workbooks.Add(MyFolder & MyFileName)
            TempFileName = ActiveWorkbook.Name
            ThisWorkbook.Sheets(1).Range("A1").Offset(i + 1, 0).Value = MyFileName
            On Error Resume Next
            Set ws = Workbooks(TempFileName).Sheets("Result")
            If ((ws Is Nothing) Or (ThisWorkbook.Sheets(3).Range("B4") = "DataRange")) Then
                For j = 1 To DataRange
                    ThisWorkbook.Sheets(1).Range("A1").Offset(i + 1, j).Value = Workbooks(TempFileName).Sheets(1).Range(tempost(j)).Value
                Next j
            Else
                Workbooks(TempFileName).Sheets("Result").Range("A1:IU1").Copy
                ThisWorkbook.Sheets(1).Activate
                ThisWorkbook.Sheets(1).Range("A1").Offset(i + 1, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues
            End If
            Set ws = Nothing
            Application.DisplayAlerts = False
            .Close False
            Application.DisplayAlerts = True
        End With
    MyFileName = Dir
    i = i + 1
    Loop
End Sub

Sub Filewriter()
    Application.ScreenUpdating = False

    MyFolder = ThisWorkbook.Sheets(3).Range("B1")
    Password = ThisWorkbook.Sheets(3).Range("B6")

    Dim i, j As Integer
    Dim tempost() As String
    DataRange = ThisWorkbook.Sheets(3).Range("B3")
    ReDim tempost(1 To DataRange)

    For j = 1 To DataRange
        tempost(j) = ThisWorkbook.Sheets(2).Cells(j, 1)
    Next j
    i = 0
    DatabaseLastRow = Application.CountA(ThisWorkbook.Sheets(1).Range("A:A"))
    For k = 1 To DatabaseLastRow - 1
        With Workbooks.Add(MyFolder & ThisWorkbook.Sheets(1).Cells(i + 2, 1))
            TempFileName = ActiveWorkbook.Name
            For j = 1 To DataRange
                Workbooks(TempFileName).Sheets(1).Range(tempost(j)).Value = ThisWorkbook.Sheets(1).Range("A1").Offset(i + 1, j).Value
            Next j
            If ThisWorkbook.Sheets(3).Range("B5") = "Lock" Then
                Workbooks(TempFileName).Sheets(1).Protect Password:=Password
            ElseIf ThisWorkbook.Sheets(3).Range("B5") = "Unlock" Then
                Workbooks(TempFileName).Sheets(1).Unprotect Password:=Password
            End If
            Application.DisplayAlerts = False
            Workbooks(TempFileName).SaveAs (MyFolder & ThisWorkbook.Sheets(1).Cells(i + 2, 1))
            .Close False
        End With
        i = i + 1
    Next k
End Sub
