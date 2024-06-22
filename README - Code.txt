
Sub StockAnalysis()

    Dim total As Double
    Dim row_index As Long
    Dim change As Double
    Dim column_index As Integer
    Dim start As Long
    Dim rowcount As Long
    Dim percent_change As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        column_index = 0
        total = 0
        change = 0
        start = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For row_index = 2 To rowcount
        
            If ws.Cells(row_index + 1, 1).Value <> ws.Cells(row_index, 1).Value Then
            
                total = total + ws.Cells(row_index, 7).Value
            
                If total = 0 Then
                
                ws.Range("I" & 2 + column_index).Value = Cells(row_index, 1).Value
                ws.Range("J" & 2 + column_index).Value = 0
                ws.Range("K" & 2 + column_index).Value = "%" & 0
                ws.Range("L" & 2 + column_index).Value = 0
                
                Else
                
                    If ws.Cells(start, 3) = 0 Then
                    
                        For find_value = start To row_index
                        
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                
                                start = find_value
                                
                                Exit For
                            
                            End If
                        
                        Next find_value
                        
                    End If
        
                    change = (ws.Cells(row_index, 6) - ws.Cells(start, 3))
                    
                    percent_change = change / ws.Cells(start, 3)
                    
                    start = row_index + 1
                    
                    ws.Range("I" & 2 + column_index) = ws.Cells(row_index, 1).Value
                    ws.Range("J" & 2 + column_index) = change
                    ws.Range("J" & 2 + column_index).NumberFormat = "0.00"
                    ws.Range("K" & 2 + column_index).Value = percent_change
                    ws.Range("K" & 2 + column_index).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + column_index).Value = total
                    
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + column_index).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + column_index).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + column_index).Interior.ColorIndex = 0
                            
                    End Select
                    
                    Select Case percent_change
                        Case Is > 0
                            ws.Range("K" & 2 + column_index).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("K" & 2 + column_index).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("K" & 2 + column_index).Interior.ColorIndex = 0
                    End Select
                    
             End If
                
                total = 0
                change = 0
                column_index = column_index + 1
                
            Else
                total = total + ws.Cells(row_index, 7).Value
            
            End If
            
        Next row_index
        
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowcount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowcount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowcount))
        
        increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
        decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
        volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)
        
        ws.Range("P2") = ws.Cells(increase + 1, 9)
        ws.Range("P3") = ws.Cells(decrease + 1, 9)
        ws.Range("P4") = ws.Cells(volume + 1, 9)
        
Next ws
End Sub









