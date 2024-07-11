'VBA Challenge Correct functioning code
Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim ticker As String
    Dim quarterOpen As Double
    Dim quarterClose As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim currentQuarter As Integer
    Dim nextQuarter As Integer
    Dim i As Long
    Dim dateValue As Variant
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Set headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        summaryRow = 2
        
        ' Determine the last row with data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables
        ticker = ws.Cells(2, 1).Value
        If IsNumeric(ws.Cells(2, 3).Value) Then
            quarterOpen = ws.Cells(2, 3).Value
        Else
            quarterOpen = 0
        End If
        totalVolume = 0
        dateValue = ws.Cells(2, 2).Value
        If IsDate(dateValue) Then
            currentQuarter = DatePart("q", dateValue)
        Else
            currentQuarter = 0
        End If
        
        ' Loop through all rows
        For i = 2 To lastRow
            dateValue = ws.Cells(i, 2).Value
            If IsDate(dateValue) Then
                nextQuarter = DatePart("q", dateValue)
            Else
                nextQuarter = 0
            End If
            
            ' Check if the ticker symbol or quarter changes
            If ws.Cells(i, 1).Value <> ticker Or nextQuarter <> currentQuarter Then
                ' Calculate quarterly change and percent change
                If IsNumeric(ws.Cells(i - 1, 6).Value) Then
                    quarterClose = ws.Cells(i - 1, 6).Value
                Else
                    quarterClose = 0
                End If
                
                quarterlyChange = quarterClose - quarterOpen
                If quarterOpen <> 0 Then
                    percentChange = (quarterlyChange / quarterOpen) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output the results to the summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                summaryRow = summaryRow + 1
                
                ' Reset variables for the next ticker or quarter
                ticker = ws.Cells(i, 1).Value
                If IsNumeric(ws.Cells(i, 3).Value) Then
                    quarterOpen = ws.Cells(i, 3).Value
                Else
                    quarterOpen = 0
                End If
                totalVolume = 0
                currentQuarter = nextQuarter
            End If
            
            ' Accumulate volume
            If IsNumeric(ws.Cells(i, 7).Value) Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Handle the last quarter for the last ticker
        If IsNumeric(ws.Cells(lastRow, 6).Value) Then
            quarterClose = ws.Cells(lastRow, 6).Value
        Else
            quarterClose = 0
        End If
        
        quarterlyChange = quarterClose - quarterOpen
        If quarterOpen <> 0 Then
            percentChange = (quarterlyChange / quarterOpen) * 100
        Else
            percentChange = 0
        End If
        ws.Cells(summaryRow, 9).Value = ticker
        ws.Cells(summaryRow, 10).Value = quarterlyChange
        ws.Cells(summaryRow, 11).Value = percentChange
        ws.Cells(summaryRow, 12).Value = totalVolume
        
        ' Format the percentage change column
        ws.Columns("K").NumberFormat = "0.00%"
    Next ws
End Sub


