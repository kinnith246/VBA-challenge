Attribute VB_Name = "Module1"
Sub ProceesQuarterlyData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openDate As Date
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim maxChange As Double
    Dim minChagnge As Double
    Dim maxTotal As Double
    
    
    For Each ws In ThisWorkbook.Worksheets
        outputRow = 2
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'create headers in new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Earliest Open Price"
        ws.Cells(1, 11).Value = "Latest Close Price"
        ws.Cells(1, 12).Value = "Quarterly Change"
        ws.Cells(1, 13).Value = "Percentage Change"
        ws.Cells(1, 14).Value = "Total Volume"
        
        'create headers for bonus task
        ws.Cells(1, 18).Value = "Ticker"
        ws.Cells(1, 19).Value = "Value"
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = "Greatest % Decrease"
        ws.Cells(4, 17).Value = "Greatest Total Volume"
        
        'add formulas for bonus task
        ws.Cells(2, 19).Formula = "=MAX(M:M)"
        ws.Cells(3, 19).Formula = "=MIN(M:M)"
        ws.Cells(4, 19).Formula = "=MAX(N:N)"
        
        ws.Cells(2, 18).Formula = "=INDEX(I:I,MATCH(S2,M:M,FALSE),1)"
        ws.Cells(3, 18).Formula = "=INDEX(I:I,MATCH(S3,M:M,FALSE),1)"
        ws.Cells(4, 18).Formula = "=INDEX(I:I,MATCH(S4,N:N,FALSE),1)"

        'set the first ticker and open price
        If lastRow > 1 Then
            ticker = ws.Cells(2, 1).Value
            openDate = ws.Cells(2, 2).Value
            openPrice = ws.Cells(2, 3).Value
        End If
        
        'loop through each row in the current sheet
        For i = 2 To lastRow
            
            'check if still same ticker name, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'set ticker
                ticker = ws.Cells(i, 1).Value
                'print ticker in summary
                ws.Cells(outputRow, 9).Value = ticker
                'set close price
                closePrice = ws.Cells(i, 6).Value
                'print close price
                ws.Cells(outputRow, 11).Value = closePrice
                'print open price
                ws.Cells(outputRow, 10).Value = openPrice
                'add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                'print total volume in summary
                ws.Cells(outputRow, 14).Value = totalVolume
                
                'calculate and print quarterly change and percentage change
                quarterChange = closePrice - openPrice
                ws.Cells(outputRow, 12).Value = quarterChange
                If openPrice <> 0 Then
                    percentChange = (quarterChange / openPrice)
                    ws.Cells(outputRow, 13).Value = percentChange
                    
                    'add formatting to change colour based on positive or negative value percentages
                    If ws.Cells(outputRow, 13).Value > 0 Then
                        ws.Cells(outputRow, 13).Interior.ColorIndex = 4 'greenfor positive
                    ElseIf ws.Cells(outputRow, 13).Value < 0 Then
                        ws.Cells(outputRow, 13).Interior.ColorIndex = 3 'red for negative
                    End If
                    
                    'add formatting to change colour based on positive or negative value quarterly change
                    If ws.Cells(outputRow, 12).Value > 0 Then
                        ws.Cells(outputRow, 12).Interior.ColorIndex = 4 'green for positive
                    ElseIf ws.Cells(outputRow, 12).Value < 0 Then
                        ws.Cells(outputRow, 12).Interior.ColorIndex = 3 'red for negative
                    End If
                    
                End If
                
                'add 1 to summary table
                outputRow = outputRow + 1
                'reset total
                totalVolume = 0
                
                'set the next ticker and open price
                If i + 1 <= lastRow Then
                    ticker = ws.Cells(i + 1, 1).Value
                    openDate = ws.Cells(i + 1, 2).Value
                    openPrice = ws.Cells(i + 1, 3).Value
                End If
            Else
                'if same ticker name then...
                If ws.Cells(i + 1, 2).Value < openDate Then
                    openDate = ws.Cells(i + 1, 2).Value
                    openPrice = ws.Cells(i + 1, 3).Value
                End If
                
                'add to the total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        'format the Percentage Change column to show percentages
        ws.Columns(13).NumberFormat = "0.00%"
        ws.Cells(2, 19).NumberFormat = "0.00%"
        ws.Cells(3, 19).NumberFormat = "0.00%"
        
        'format relevant columns and cells to avoid scientific notation
        ws.Columns(14).NumberFormat = "0"    'total volume
        ws.Cells(4, 19).NumberFormat = "0"   'greatest total volume
        
    Next ws
End Sub
