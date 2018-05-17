Sub Stock_Market()

Dim ws As Worksheet

Dim Ticker, TickerGreatestTotalValue, TickerGreatestPercentIncrease, TickerGreatestPercentDecrease As String
Dim TotalStockVolume As Double


Dim SummaryTableRow As Integer
Dim OpenPrice, ClosePrice, YearlyChange, PercentChange, GreatestPercentIncrease, GreatestPercentDecrease, GreatestTotalValue As Double

Dim LastRow As Long

For Each ws In Worksheets
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly Change"
     ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
     SummaryTableRow = 2
     TotalStockVolume = 0
     GreatestTotalValue = 0
     GreatestPercentIncrease = 0
     GreatestPercentDecrease = 0
     TickerGreatestPercentIncrease = " "
     TickerGreatestTotalValue = " "
     TickerGreatestPercentDecrease = " "
     OpenPrice = ws.Cells(2, 3).Value
     For I = 2 To LastRow
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
              Ticker = ws.Cells(I, 1).Value
              TotalStockVolume = TotalStockVolume + ws.Cells(I, 7).Value
              ws.Range("I" & SummaryTableRow).Value = Ticker
              ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
              
             If TotalStockVolume > GreatestTotalValue Then
                 GreatestTotalValue = TotalStockVolume
                 TickerGreatestTotalValue = Ticker
             End If
             
              TotalStockVolume = 0
              ClosePrice = ws.Cells(I, 6)
              YearlyChange = ClosePrice - OpenPrice
              If OpenPrice <> 0 Then
                  PercentChange = YearlyChange / OpenPrice
              Else
                 PercentChange = 0
              End If
              If PercentChange > GreatestPercentIncrease Then
                 GreatestPercentIncrease = PercentChange
                 TickerGreatestPercentIncrease = Ticker
              ElseIf PercentChange <= GreatestPercentDecrease Then
                 GreatestPercentDecrease = PercentChange
                 TickerGreatestPercentDecrease = Ticker
              End If
            
              ws.Range("J" & SummaryTableRow).Value = YearlyChange
              ws.Range("K" & SummaryTableRow).Value = PercentChange
              ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
              If YearlyChange >= 0 Then
                 ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
              Else
                 ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
              End If
              SummaryTableRow = SummaryTableRow + 1
              OpenPrice = ws.Cells(I + 1, 3).Value
              ClosePrice = 0
         Else
              TotalStockVolume = TotalStockVolume + ws.Cells(I, 7).Value
         End If
     Next I




ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(2, 16).Value = TickerGreatestPercentIncrease
ws.Cells(2, 17).Value = GreatestPercentIncrease
ws.Cells(2, 17).NumberFormat = "0.00%"

ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(3, 16).Value = TickerGreatestPercentDecrease
ws.Cells(3, 17).Value = GreatestPercentDecrease
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Cells(4, 15).Value = "Greatest total volume"
ws.Cells(4, 16).Value = TickerGreatestTotalValue
ws.Cells(4, 17).Value = GreatestTotalValue


Next ws

End Sub

