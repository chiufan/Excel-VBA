Attribute VB_Name = "Module1"
Sub ReadData(StockNo As String)
Columns("A:D").Clear
webURL = "URL;http://www.aastocks.com/tc/ltp/RTQuoteContent.aspx?symbol=" & StockNo & "&process=y"

With ActiveSheet.QueryTables.Add(Connection:=webURL, Destination:=Range("A1"))
    .RefreshStyle = xlOverwriteCells
    .WebTables = "3"
    .Refresh BackgroundQuery:=False
End With

Range("A8").Value = "最後更新: " & Now()
ThisWorkbook.Names(1).Delete
End Sub
