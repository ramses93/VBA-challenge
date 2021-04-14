Sub Main()
    Dim ws As Worksheet
    
    For Each ws In Sheets
    
        AnalyzeSheet ws
    
    Next ws
    
End Sub

Sub AnalyzeSheet(ws As Worksheet)
    
    ' Row index of current working calulations
    Dim row_index As Long
    row_index = 2
    
    ' Ticker data keepers
    Dim ticker_from As Long
    ticker_from = 2
    
    Dim ticker_to As Long
    ticker_to = 2
    
    Dim ticker_volume As Double
    ticker_volume = 0
    
    
    ' creating column headers
    ws.Cells(1, 9) = "ticker"
    ws.Cells(1, 10) = "yearly_change"
    ws.Cells(1, 11) = "percent_change"
    ws.Cells(1, 12) = "Total_volume"
    
    ws.Cells(1, 16) = "ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "greatest_percent_increase"
    ws.Cells(3, 15) = "greatest_percent_decrease"
    ws.Cells(4, 15) = "greates_total_Volume"
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    Dim last_row_index As Long
    last_row_index = ws.UsedRange.Rows.Count
    
    For I = 2 To last_row_index
    
        ticker_volume = ws.Cells(I, 7).Value + ticker_volume
        
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            ws.Cells(row_index, 9) = ws.Cells(I, 1).Value

            ticker_to = CLng(I)
            
            CalculateYearlyChange ws, ticker_from, ticker_to, row_index
            CalculatePercentChange ws, ticker_from, ticker_to, row_index
            
            ws.Cells(row_index, 12) = ticker_volume
            
            ' Update the ticker details for the next ticker
            row_index = row_index + 1
            ticker_from = CLng(I) + 1
            ticker_volume = 0
        End If
            
    Next I
    
    ' Starting additional analysis
    Dim last_row_ticker_overview As Long
    last_row_ticker_overview = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    
    For I = 2 To last_row_ticker_overview
        If ws.Cells(I, 11) > ws.Cells(2, 17) Then
            ws.Cells(2, 16) = ws.Cells(I, 9)
            ws.Cells(2, 17) = ws.Cells(I, 11)
        End If
        
        If ws.Cells(I, 11) < ws.Cells(3, 17) Then
            ws.Cells(3, 16) = ws.Cells(I, 9)
            ws.Cells(3, 17) = ws.Cells(I, 11)
        End If
        
        If ws.Cells(I, 12) > ws.Cells(4, 17) Then
            ws.Cells(4, 16) = ws.Cells(I, 9)
            ws.Cells(4, 17) = ws.Cells(I, 12)
        End If
        
    Next I
    
End Sub

Sub CalculateYearlyChange(ws As Worksheet, from_row As Long, to_row As Long, row_index As Long)
    ws.Cells(row_index, 10) = ws.Cells(to_row, 6).Value - ws.Cells(from_row, 3).Value
    
    If ws.Cells(row_index, 10).Value > 0 Then
        ws.Cells(row_index, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(row_index, 10).Value = 0 Then
        ws.Cells(row_index, 10).Interior.ColorIndex = 2
    Else
        ws.Cells(row_index, 10).Interior.ColorIndex = 3
    End If
    
End Sub

Sub CalculatePercentChange(ws As Worksheet, from_row As Long, to_row As Long, row_index As Long)
    ws.Cells(row_index, 11) = ((ws.Cells(to_row, 6).Value - ws.Cells(from_row, 3).Value) / ws.Cells(from_row, 3).Value)
    ws.Cells(row_index, 11).NumberFormat = "0.00%"
End Sub


