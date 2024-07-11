'Calculate Extreme Values
Sub CalculateExtremeValues()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Range, percentChangeColumn As Range, volumeColumn As Range
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim increaseTicker As String, decreaseTicker As String, volumeTicker As String
    Dim i As Long
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column I (Ticker column)
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        ' Define the columns
        Set tickerColumn = ws.Range("I2:I" & lastRow)
        Set percentChangeColumn = ws.Range("K2:K" & lastRow)
        Set volumeColumn = ws.Range("L2:L" & lastRow)
        
        ' Initialize variables
        greatestIncrease = -1 ' Initialize to a very low value
        greatestDecrease = 1  ' Initialize to a very high value
        greatestVolume = 0
        increaseTicker = ""
        decreaseTicker = ""
        volumeTicker = ""
        
        ' Loop through the data to find the extreme values and corresponding tickers
        For i = 2 To lastRow
            ' Check for greatest increase
            If ws.Cells(i, "K").Value > greatestIncrease Then
                greatestIncrease = ws.Cells(i, "K").Value
                increaseTicker = ws.Cells(i, "I").Value
            End If
            
            ' Check for greatest decrease
            If ws.Cells(i, "K").Value < greatestDecrease Then
                greatestDecrease = ws.Cells(i, "K").Value
                decreaseTicker = ws.Cells(i, "I").Value
            End If
            
            ' Check for greatest total volume
            If ws.Cells(i, "L").Value > greatestVolume Then
                greatestVolume = ws.Cells(i, "L").Value
                volumeTicker = ws.Cells(i, "I").Value
            End If
        Next i
        
        ' Output the results
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 16).Value = increaseTicker
        ws.Cells(2, 17).Value = greatestIncrease * 0.01 ' Move decimal point
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = decreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease * 0.01 ' Move decimal point
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = volumeTicker
        ws.Cells(4, 17).Value = greatestVolume
        ws.Cells(4, 17).NumberFormat = "#,##0"
        
        ' Auto-fit columns for better readability
        ws.Columns("O:Q").AutoFit
    Next ws
    
    MsgBox "Extreme values calculated and displayed for all worksheets!"
End Sub
