Sub tickerSummary()

Dim ws As Worksheet
Dim ticker As Integer
Dim yearClose, yearOpen, openCount, yearChange, perChange, stock_total, max_num, min_num, max_stock As Double

For Each ws In Worksheets

last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

row_out = 2
max_num = 0
min_num = 0
max_stock = 0
    
'Titles
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'CHALLENGE Titles
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

    For i = 2 To last_row
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Ticker
            ws.Cells(row_out, 9).Value = ws.Cells(i, 1).Value
            
            'Find yearOpen and yearClose; Yearly Change
            yearClose = ws.Cells(i, 6).Value
            yearOpen = ws.Cells((i - openCount), 3).Value

            yearChange = yearClose - yearOpen
            ws.Cells(row_out, 10).Value = yearChange
            ws.Cells(row_out, 10).NumberFormat = "$0.00"
            
            'Conditional Formatting
            If yearChange > 0 Then
                ws.Cells(row_out, 10).Interior.ColorIndex = 4 'Green
            ElseIf yearChange < 0 Then
                ws.Cells(row_out, 10).Interior.ColorIndex = 3 'Red
            Else
                ws.Cells(row_out, 10).Interior.ColorIndex = 0 'No Fill
            End If
            
            '% Change
            If yearOpen = 0 Then
                ws.Cells(row_out, 11).Value = 0
            Else
                perChange = yearChange / yearOpen
                ws.Cells(row_out, 11).Value = perChange
            End If
            ws.Cells(row_out, 11).NumberFormat = "0.00%"
            
            'Total Stock Volume
            stock_total = stock_total + ws.Cells(i, 7).Value
            ws.Cells(row_out, 12).Value = stock_total
            ws.Cells(row_out, 12).NumberFormat = "#,##"
            
            'CHALLENGE : Greatest Total Volume
            If stock_total > max_stock Then
                max_stock = stock_total
                ws.Cells(4, 16).Value = max_stock
                ws.Cells(4, 16).NumberFormat = "#,##"
                ws.Cells(4, 15).Value = ws.Cells(i, 1).Value                
            End If
            
            'Row Counter
            row_out = row_out + 1
            
            'Reset
            stock_total = 0
            openCount = 0
            
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            stock_total = stock_total + ws.Cells(i, 7).Value
            openCount = openCount + 1
        End If
        
        'CHALLENGES : Greatest % Increase & Decrease
        If perChange > max_num Then 'Greatest % Increase
            max_num = perChange
            ws.Cells(2, 16).Value = max_num
            ws.Cells(2, 16).NumberFormat = "0.00%"
            ws.Cells(2, 15).Value = ws.Cells(i, 1).Value
        ElseIf perChange < min_num Then 'Greatest % Decrease
            min_num = perChange
            ws.Cells(3, 16).Value = min_num
            ws.Cells(3, 16).NumberFormat = "0.00%"
            ws.Cells(3, 15).Value = ws.Cells(i, 1).Value
        End If
    Next i    
Next
End Sub