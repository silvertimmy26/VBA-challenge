Attribute VB_Name = "Module1"
Sub StockData()

    Dim ws As Worksheet
    For Each ws In Worksheets
    
        ' Variables for first part of output
        Dim NumRows As Long, i As Long
        Dim stockVolume As Double ' Tried as a long, but it needs to be double to avoid overflow
        Dim currentTicker As String
        Dim firstPrice As Double, lastPrice As Double
        Dim tickerCount As Integer ' Tracks where to put new ticker symbols
        ' Second part of the output
        Dim maxPercentIncrease As Double, maxPercentDecrease As Double, maxVolume As Double
        Dim maxIncreaseRow As Long, maxDecreaseRow As Long, maxVolumeRow As Long
        Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String
        
        ' Count rows
        NumRows = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.Count
    
        ' Setting headers for output
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").ColumnWidth = 15
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").ColumnWidth = 14
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").ColumnWidth = 17
        ws.Range("L1").Value = "Total Stock Volume"
    
        ' Initializes first ticker and prices
        currentTicker = ws.Range("A2").Value
        firstPrice = ws.Range("C2").Value
        ws.Range("I2").Value = currentTicker ' First Ticker output
        stockVolume = 0
        tickerCount = 2
    
        For i = 2 To NumRows
        
            stockVolume = stockVolume + ws.Range("G" & i).Value
    
            ' If the next row is a new ticker
            If ws.Range("A" & i + 1).Value <> currentTicker Then
            
                lastPrice = ws.Range("F" & i).Value
    
                ' Sets outputs for current ticker
                ws.Range("J" & tickerCount).Value = lastPrice - firstPrice ' Quarterly change
                ws.Range("K" & tickerCount).Value = ((lastPrice - firstPrice) / firstPrice) ' Percent change
                ws.Range("L" & tickerCount).Value = stockVolume ' Total stock volume
    
                ' Next ticker if we're not at the end
                If i < NumRows Then
                    tickerCount = tickerCount + 1
                    currentTicker = ws.Range("A" & i + 1).Value
                    firstPrice = ws.Range("C" & i + 1).Value
                    ws.Range("I" & tickerCount).Value = currentTicker ' Ticker
                    stockVolume = 0
                End If
                
            End If
            
        Next i
        
        ' Next section of output, greatest % changes and total volume
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("O4").ColumnWidth = 20
        
        maxPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))  ' Greatest % Increase Value
        maxIncreaseRow = Application.WorksheetFunction.Match(maxPercentIncrease, ws.Range("K:K"), 0) ' Row with max % increase
        maxIncreaseTicker = ws.Range("I" & maxIncreaseRow).Value ' Greatest % Increase Ticker
        
        maxPercentDecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))  ' Greatest % Decrease Value
        maxDecreaseRow = Application.WorksheetFunction.Match(maxPercentDecrease, ws.Range("K:K"), 0) ' Row with max % decrease
        maxDecreaseTicker = ws.Range("I" & maxDecreaseRow).Value ' Greatest % Decrease Ticker
        
        maxVolume = Application.WorksheetFunction.Max(ws.Range("L:L")) ' Greatest Total Volume Value
        maxVolumeRow = Application.WorksheetFunction.Match(maxVolume, ws.Range("L:L"), 0) ' Row with max volume
        maxVolumeTicker = ws.Range("I" & maxVolumeRow).Value ' Greatest Total Volume Ticker
        
        ws.Range("P2").Value = maxIncreaseTicker
        ws.Range("Q2").Value = maxPercentIncrease
        ws.Range("P3").Value = maxDecreaseTicker
        ws.Range("Q3").Value = maxPercentDecrease
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q4").Value = maxVolume
        
        ' We want these cells to be percentages to two decimal places
        ws.Range("K2:K" & tickerCount).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ' Range("K2:K" & tickerCount).NumberFormat = "General"
        ' Range("Q2:Q3").NumberFormat = "General"

    Next ws

End Sub


