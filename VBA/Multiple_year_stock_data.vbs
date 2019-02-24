Sub StockData()

    'Loop through all worksheets
    For Each ws In Worksheets

        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Add headers for summary tables
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        'Create variable to hold total volume for each stock
        Dim VolumeTotal As Double
        VolumeTotal = 0
  
        'Create variable to hold opening price for each stock
        Dim OpenPrice As Double
        OpenPrice = 0
  
        'Create variable to hold the lowest % ticker ID
        Dim LowTicker As String
  
        'Create variable to hold the lowest percentage
        Dim LowPercent As Double
        LowPercent = 0
  
        'Create variable to hold the highest % ticker ID
        Dim HighTicker As String
    
        'Create variable to hold the highest percentage
        Dim HighPercent As Double
        HighPercent = 0
  
        'Create variable to hold the highest volume
        Dim HighVolume As Double
        HighVolume = 0
  
        'Create variable to hold the highest volume ticker ID
        Dim HighVolTicker As String
    
        'Create variable to count stock number for putting ticker and total in proper row
        Dim StockNumber As Integer
        StockNumber = 1
  
        'Loop through rows in the column
        For i = 2 To LastRow
        VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
        OpenPrice = ws.Cells(i, 3).Value
        
            'Detects when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or ws.Cells(i + 1, 1).Value = "" Then

                'Adds summary data to table
                StockNumber = StockNumber + 1
                ws.Cells(StockNumber, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(StockNumber, 10).Value = ws.Cells(StockNumber, 6).Value - OpenPrice
            
                'Color cell green if positive or red if negative
                If ws.Cells(StockNumber, 10).Value > 0 Then
                    ws.Cells(StockNumber, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(StockNumber, 10).Value < 0 Then
                    ws.Cells(StockNumber, 10).Interior.ColorIndex = 3
                End If
            
                'Avoid divide-by-zero error if opening price is 0.00
                If OpenPrice <> 0 Then
                ws.Cells(StockNumber, 11).Value = ((ws.Cells(StockNumber, 6).Value / OpenPrice - 1) * 100) & "%"
                Else
                    ws.Cells(StockNumber, 11).Value = ((ws.Cells(StockNumber, 6).Value / 0.000000001 - 1) * 100) & "%"
                End If
            
                ws.Cells(StockNumber, 12).Value = VolumeTotal
            
                'Record ticker ID and percentage if higher or lower than current value
                If ws.Cells(StockNumber, 11).Value < LowPercent Then
                    LowPercent = ws.Cells(StockNumber, 11)
                    LowTicker = ws.Cells(i, 1).Value
                ElseIf ws.Cells(StockNumber, 11).Value > HighPercent Then
                    HighPercent = ws.Cells(StockNumber, 11)
                    HighTicker = ws.Cells(i, 1).Value
                End If
            
                'Record volume if higher than current value
                If ws.Cells(StockNumber, 12).Value > HighVolume Then
                    HighVolume = ws.Cells(StockNumber, 12)
                    HighVolTicker = ws.Cells(i, 1).Value
                End If
            
                'Adjust columns to fit content
                ws.Columns("I:L").AutoFit
                ws.Columns("O:Q").AutoFit
            
                'Reset variable(s)
                VolumeTotal = 0

            End If

        Next i
        
        'Add data to second summary table
        ws.Range("P2").Value = HighTicker
        ws.Range("Q2").Value = HighPercent * 100 & "%"
        ws.Range("P3").Value = LowTicker
        ws.Range("Q3").Value = LowPercent * 100 & "%"
        ws.Range("P4").Value = HighVolTicker
        ws.Range("Q4").Value = HighVolume

    Next ws

End Sub

