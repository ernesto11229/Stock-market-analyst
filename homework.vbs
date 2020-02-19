Sub loopTest()
Dim TickerName As String
Dim StockTotal As Variant
StockTotal = 0
Dim SummaryTable As Variant
SummaryTable = 2
Dim i As Long
Dim mylastrow As Long

For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


mylastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To mylastrow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

TickerName = ws.Cells(i, 1).Value

StockTotal = StockTotal + ws.Cells(i, 7).Value

Range("i" & SummaryTable).Value = TickerName

Range("l" & SummaryTable).Value = StockTotal

SummaryTable = SummaryTable + 1

StockTotal = 0

Else

StockTotal = StockTotal + ws.Cells(i, 7).Value


End If


Next i

Next ws

End Sub

