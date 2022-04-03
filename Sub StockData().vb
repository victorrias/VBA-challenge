Sub StockData()


'Bonus - Make the appropriate adjustments to your VBA script to allow it to run on every worksheet
Dim ws As Worksheet
For Each ws In Worksheets


'Create a script that loops through all of stocks for one year and outputs the:
'Ticker symbol

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


Dim LastRowDetail As Long

LastRowDetail = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("A2:A" & LastRowDetail).Copy ws.Range("I2")

ws.Range("I1:I" & LastRowDetail).RemoveDuplicates Columns:=1, Header:=xlYes


'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year

Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim LastRowSummary As Integer

LastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To LastRowDetail
For j = 2 To LastRowSummary


If (ws.Cells(i, 1).Value = ws.Cells(j, 9).Value) And (ws.Cells(i, 2).Value = "20180102") Or (ws.Cells(i, 2).Value = "20190102") Or (ws.Cells(i, 2).Value = "20200102") Then
OpeningPrice = ws.Cells(i, 3).Value
End If

If (ws.Cells(i, 1).Value = ws.Cells(j, 9).Value) And (ws.Cells(i, 2).Value = "20181231") Or (ws.Cells(i, 2).Value = "20191231") Or (ws.Cells(i, 2).Value = "20201231") Then
ClosingPrice = ws.Cells(i, 6).Value
End If

If (ws.Cells(i, 1).Value = ws.Cells(j, 9).Value) Then
ws.Cells(j, 10).Value = ClosingPrice - OpeningPrice
End If


'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

If (ws.Cells(i, 1).Value = ws.Cells(j, 9).Value) Then
ws.Cells(j, 11).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
End If


'The total stock volume of the stock

Dim TotalStockVolume As LongLong

TotalStockVolume = 0

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And (ws.Cells(i, 1).Value = ws.Cells(j, 9).Value) Then
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
ws.Cells(j, 12).Value = TotalStockVolume
TotalStockVolume = 0

Else
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
End If


'Conditional Formatting to Yearly Change
If (ws.Cells(j, 10).Value > 0) Then
ws.Cells(j, 10).Interior.ColorIndex = 4

ElseIf (ws.Cells(j, 10).Value < 0) Then
ws.Cells(j, 10).Interior.ColorIndex = 3

End If



Next j

Next i

Next ws


End Sub