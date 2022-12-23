Attribute VB_Name = "Module1"
Sub StockData():

'Part 1

    For Each WS In Worksheets
    
    'Creating headers for the worksheet'
    
        WS.Range("I1").Value = "Ticker"
        WS.Range("j1").Value = "Yearly Change"
        WS.Range("k1").Value = "Percent Change"
        WS.Range("l1").Value = "Total Stock Volume"
        WS.Range("p1").Value = "Ticker"
        WS.Range("q1").Value = "Value"
        WS.Range("o2").Value = "Greatest % Increase"
        WS.Range("o3").Value = "Greatest % Decrease"
        WS.Range("o4").Value = "Greatest Total Volume"
    
    'Creating variables'
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickerCount As Long
        Dim LastRowA As Long
        Dim LastRowB As Long
        Dim PerChange As Double
        Dim Volume As Double
        Dim GreatIncrease As Double
        Dim GreatDecrease As Double
        Dim GreatVolume As Double
   
        WorksheetName = WS.Name

        TickerCount = 2
 
        j = 2
        
        'Find the last cell'
        LastRowA = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
            
            For i = 2 To LastRowA
            
                If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                WS.Cells(TickerCount, 9).Value = WS.Cells(i, 1).Value
                
                'Calculate and write Yearly Change
                WS.Cells(TickerCount, 10).Value = WS.Cells(i, 6).Value - WS.Cells(j, 3).Value
                
                    If WS.Cells(TickerCount, 10).Value < 0 Then
                    'Red background color
                    WS.Cells(TickerCount, 10).Interior.ColorIndex = 3
                    Else
                    'Green background color
                    WS.Cells(TickerCount, 10).Interior.ColorIndex = 4
                    End If
                    
                    'Calculate percent change
                    If WS.Cells(j, 3).Value <> 0 Then
                    PerChange = ((WS.Cells(i, 6).Value - WS.Cells(j, 3).Value) / WS.Cells(j, 3).Value)
                    WS.Cells(TickerCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    WS.Cells(TickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate total volume
                WS.Cells(TickerCount, 12).Value = Application.WorksheetFunction.Sum(WS.Range(WS.Cells(j, 7), WS.Cells(i, 7)))
                
                
                'Increase TickerCount
                TickerCount = TickerCount + 1
                
                'Start new row
                j = i + 1
                
                End If
            
            Next i
            
            
'Part 2
        LastRowB = WS.Cells(Rows.Count, 9).End(xlUp).Row
        'Loop and find the greatest and the lowest
        For x = 2 To LastRowB
            If WS.Cells(x, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & LastRowB)) Then
                WS.Cells(2, 16).Value = Cells(x, 9).Value
                WS.Cells(2, 17).Value = Cells(x, 11).Value
                WS.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & LastRowB)) Then
                WS.Cells(3, 16).Value = Cells(x, 9).Value
                WS.Cells(3, 17).Value = Cells(x, 11).Value
                WS.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & LastRowB)) Then
                WS.Cells(4, 16).Value = Cells(x, 9).Value
                WS.Cells(4, 17).Value = Cells(x, 12).Value
            End If
        Next x

            
    Next WS
        
End Sub



