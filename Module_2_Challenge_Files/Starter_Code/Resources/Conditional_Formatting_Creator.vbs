'Conditional Formatting Creator AKA Picasso
Sub Picasso()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim quarterlyChangeRange As Range
    Dim percentChangeRange As Range
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column I (Ticker column)
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        ' Define the ranges for quarterly change and percent change
        Set quarterlyChangeRange = ws.Range("J2:J" & lastRow)
        Set percentChangeRange = ws.Range("K2:K" & lastRow)
        
        ' Clear existing conditional formatting
        quarterlyChangeRange.FormatConditions.Delete
        percentChangeRange.FormatConditions.Delete
        
        ' Apply conditional formatting to quarterly change column
        With quarterlyChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(0, 255, 0) ' Green for positive change
        End With
        
        With quarterlyChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
            .Interior.Color = RGB(255, 0, 0) ' Red for negative change
        End With
        
        ' Apply conditional formatting to percent change column
        With percentChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(0, 255, 0) ' Green for positive change
        End With
        
        With percentChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
            .Interior.Color = RGB(255, 0, 0) ' Red for negative change
        End With
    Next ws
    
    MsgBox "Conditional formatting applied successfully to all worksheets!"
End Sub
