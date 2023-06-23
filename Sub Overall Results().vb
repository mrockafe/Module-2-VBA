Sub Ticker_Groups()
    Dim lastRow As Long
    Dim tickerRange As Range
    Dim tickerCell As Range
    Dim ticker As String
    Dim firstValue As Double
    Dim lastValue As Double
    Dim sumValue As Double
    Dim percentChange As Double
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim resultsRow As Long
    
    ' Find the last row in column 1
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Create results table
    resultsRow = 2
    Cells(resultsRow, 9).Value = "Ticker Group"
    Cells(resultsRow, 10).Value = "Subtraction Result"
    Cells(resultsRow, 11).Value = "Percent Change"
    Cells(resultsRow, 12).Value = "Sum of Column 7"
    
    ' Loop through each row
    For I = 2 To lastRow
        ' Check if the ticker name changes
        If Cells(I, 1).Value <> ticker Then
            ' Calculate and then subtract the difference between the open and close
            If Not tickerRange Is Nothing Then
                firstValue = Cells(tickerRange.Cells(1).Row, 3).Value
                lastValue = Cells(tickerRange.Cells(tickerRange.Cells.Count).Row, 6).Value
                sumValue = Application.WorksheetFunction.Sum(Range("G" & tickerRange.Cells(1).Row & ":G" & tickerRange.Cells(tickerRange.Cells.Count).Row))
                percentChange = (lastValue - firstValue) / firstValue
                
                Cells(resultsRow, 9).Value = ticker
                Cells(resultsRow, 10).Value = lastValue - firstValue
                Cells(resultsRow, 11).Value = Format(percentChange, "0.00%")
                Cells(resultsRow, 12).Value = sumValue
                
                ' Update results table for sheet
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxPercentIncreaseTicker = ticker
                ElseIf percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxPercentDecreaseTicker = ticker
                End If
                
                If sumValue > maxVolume Then
                    maxVolume = sumValue
                    maxVolumeTicker = ticker
                End If
                
                resultsRow = resultsRow + 1
            End If
            
            ' Move to the next same grouped tickers
            ticker = Cells(I, 1).Value
            Set tickerRange = Range("A" & I)
        Else
            'set range for the current ticker group
            Set tickerRange = Union(tickerRange, Range("A" & I))
        End If
    Next I

    ' Generate overall results table
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(2, 16).Value = maxPercentIncreaseTicker
    Cells(3, 16).Value = maxPercentDecreaseTicker
    Cells(4, 16).Value = maxVolumeTicker
    Cells(2, 17).Value = Format(maxPercentIncrease, "0.00%")
    Cells(3, 17).Value = Format(maxPercentDecrease, "0.00%")
    Cells(4, 17).Value = maxVolume
    
    End Sub