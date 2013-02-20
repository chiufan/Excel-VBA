Sub ReadData(StockNo As String)
    Columns("A:D").Clear
    webURL = "URL;http://www.aastocks.com/tc/ltp/RTQuoteContent.aspx?symbol=" & StockNo & "&process=y"
    
    With ActiveSheet.QueryTables.Add(Connection:=webURL, Destination:=Range("A1"))
        .RefreshStyle = xlOverwriteCells
        .WebTables = "3"
        .Refresh BackgroundQuery:=False
    End With
    
    With ThisWorkbook.Sheets(1).Range("C3:D7")
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        With .FormatConditions(1).Font
            .Color = -16752384
        End With
        With .FormatConditions(1).Interior
            .Color = 13561798
        End With
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        With .FormatConditions(2).Font
            .Color = -16383844
        End With
        With .FormatConditions(2).Interior
            .Color = 13551615
        End With
    End With
    
    Range("A8").Value = "最後更新: " & Now()
    ThisWorkbook.Names(1).Delete
End Sub
