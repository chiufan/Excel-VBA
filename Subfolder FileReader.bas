Sub FileReader()
    
    
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(1).Range("A2:IV65536").Clear
    MyFolder = ThisWorkbook.Worksheets(3).Range("B1")
    MyFileName = Dir(MyFolder & ThisWorkbook.Worksheets(3).Range("B2"))

    Dim i, j As Integer
    Dim tempost() As String
    DataRange = ThisWorkbook.Worksheets(3).Range("B3")
    ReDim tempost(1 To DataRange)
    For j = 1 To DataRange
        tempost(j) = ThisWorkbook.Worksheets(2).Cells(j, 1)
    Next j
    
    
    If Application.Version > 11 Then
        i = 0
        Do Until MyFileName = ""
            With Workbooks.Add(MyFolder & MyFileName)
                TempFileName = ActiveWorkbook.Name
                ThisWorkbook.Worksheets(1).Range("A1").Offset(i + 1, 0).Value = MyFolder & MyFileName
                ThisWorkbook.Worksheets(1).Range("A1").Offset(i + 1, 1).Value = MyFileName
                For j = 1 To DataRange
                    ThisWorkbook.Worksheets(1).Range("A1").Offset(i + 1, j + 1).Value = Workbooks(TempFileName).Worksheets(1).Range(tempost(j)).Value
                Next j
                .Close False
            End With
        MyFileName = Dir
        i = i + 1
        Loop
        
    Else

        Dim fs As FileSearch
        Set fs = Application.FileSearch
        
        
        With fs
            .SearchSubFolders = True ' set to true if you want sub-folders included
            .FileType = msoFileTypeExcelWorkbooks 'can modify to just Excel files eg with msoFileTypeExcelWorkbooks
            .LookIn = MyFolder 'modify this to where you want to serach
            If .Execute > 0 Then
            
            
                For i = 1 To .FoundFiles.Count
                TempPath = .FoundFiles(i)
                With Workbooks.Add(.FoundFiles(i))
                    TempFileName = ActiveWorkbook.Name
                    ThisWorkbook.Worksheets(1).Range("A1").Offset(i, 0).Value = TempPath 'Myfolder & TempFileName & ".xls"
                    ThisWorkbook.Worksheets(1).Range("A1").Offset(i, 1).Value = TempFileName & ".xls"
                    For j = 1 To DataRange
                        ThisWorkbook.Worksheets(1).Range("A1").Offset(i, j + 1).Value = Workbooks(TempFileName).Worksheets(1).Range(tempost(j)).Value
                    Next j
                    .Close False
                End With
                Next
            Else
            End If
        End With
    End If
    
End Sub

Sub Filewriter()
    Application.ScreenUpdating = False

    MyFolder = ThisWorkbook.Worksheets(3).Range("B1")
    MyFileName = Dir(MyFolder & ThisWorkbook.Worksheets(3).Range("B2"))

    Dim i, j As Integer
    Dim tempost() As String
    DataRange = ThisWorkbook.Worksheets(3).Range("B3")
    ReDim tempost(1 To DataRange)

    For j = 1 To DataRange
        tempost(j) = ThisWorkbook.Worksheets(2).Cells(j, 1)
    Next j
    
    
    
    If Application.Version > 11 Then
    
        numfile = Application.CountA(ThisWorkbook.Worksheets(1).Range("A:A"))
    
    For i = 0 To (numfile - 2)
        With Workbooks.Add(MyFolder & ThisWorkbook.Worksheets(1).Cells(i + 2, 2))
            TempFileName = ActiveWorkbook.Name
            For j = 1 To DataRange
                Workbooks(TempFileName).Worksheets(1).Range(tempost(j)).Value = ThisWorkbook.Worksheets(1).Range("A1").Offset(i + 1, j + 1).Value
            Next j
            Application.DisplayAlerts = False
            Workbooks(TempFileName).SaveAs (MyFolder & ThisWorkbook.Worksheets(1).Cells(i + 2, 2))
            .Close False
        End With
    Next i
    Else
        Dim fs As FileSearch
        Set fs = Application.FileSearch
        
        
        With fs
            .SearchSubFolders = True ' set to true if you want sub-folders included
            .FileType = msoFileTypeExcelWorkbooks 'can modify to just Excel files eg with msoFileTypeExcelWorkbooks
            .LookIn = MyFolder 'modify this to where you want to serach
            If .Execute > 0 Then
            
            
                For i = 1 To (Application.CountA(ThisWorkbook.Worksheets(1).Range("A:A")) - 1)
                TempPath = .FoundFiles(i)
                With Workbooks.Add(.FoundFiles(i))
                    TempFileName = ActiveWorkbook.Name
                    For j = 1 To DataRange
                        Workbooks(TempFileName).Worksheets(1).Range(tempost(j)).Value = ThisWorkbook.Worksheets(1).Range("A1").Offset(i, j + 1).Value
                    Next j
                    Application.DisplayAlerts = False
                Workbooks(TempFileName).SaveAs (ThisWorkbook.Worksheets(1).Cells(i + 1, 1))
                    .Close False
                End With
                Next
            Else
            End If
        End With
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
End Sub
