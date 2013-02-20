Sub ReadData(StockNo As String)
Columns("A:D").Clear
webURL = "URL;http://www.aastocks.com/tc/ltp/RTQuoteContent.aspx?symbol=" & StockNo & "&process=y"

With ActiveSheet.QueryTables.Add(Connection:=webURL, Destination:=Range("A1"))
    .RefreshStyle = xlOverwriteCells
    .WebTables = "3"
    .Refresh BackgroundQuery:=False
End With

'StockData = ThisWorkbook.Sheets(1).Range("B3:D7")
ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
    Formula1:="=0"
ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions(ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions.Count).SetFirstPriority
With ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions(1).Font
    .Color = -16752384
    .TintAndShade = 0
End With
With ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
End With
ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions(1).StopIfTrue = False
ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
    Formula1:="=0"
ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions(ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions.Count).SetFirstPriority
With ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions(1).Font
    .Color = -16383844
    .TintAndShade = 0
End With
With ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
End With
ThisWorkbook.Sheets(1).Range("B3:D7").FormatConditions(1).StopIfTrue = False

Range("A8").Value = "最後更新: " & Now()
ThisWorkbook.Names(1).Delete
End Sub
