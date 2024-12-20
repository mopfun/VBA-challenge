Attribute VB_Name = "Module1"
Sub QuarterlyStocks()
   
    ' Clarify the worksheet
    Dim ws As Worksheet
    
    ' Extracting/consolidating all items in new table
    Dim i As Long
    Dim ticker_name As String
    Dim qrtlyChange As Double
    Dim percentChange As Double
    Dim stock_vol As Double
    
    Dim openPrice As Double
    Dim closePrice As Double

    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVol As Double
    Dim maxTicker As String
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim totalVolTicker As String
    
    For Each ws In Worksheets
        
        'reset marks
        stock_vol = 0
        qrtlyChange = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVol = 0
        
        Dim new_table As Long
        new_table = 2
        
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        openPrice = ws.Cells(2, 3).Value 'rest open price
    
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then
                ticker_name = ws.Cells(i, 1).Value ' Setting the new ticker name
                closePrice = ws.Cells(i, 6).Value
                stock_vol = stock_vol + ws.Cells(i, 7).Value ' Setting the stock vol
                
                'openPrice = ws.Cells(2, 3).Value
                qrtlyChange = closePrice - openPrice
                
                
                If openPrice <> 0 Then
                    percentChange = (closePrice - openPrice) / openPrice
                Else
                    percentChange = 0
                End If
    
                ' New table outputs
                ws.Range("I" & new_table).Value = ticker_name
                ws.Range("J" & new_table).Value = qrtlyChange
                ws.Range("K" & new_table).Value = percentChange
                ws.Range("K" & new_table).NumberFormat = "0.00%" 'percentage display
                ws.Range("L" & new_table).Value = stock_vol
                
                ' Cell color
                If qrtlyChange > 0 Then
                ws.Range("J" & new_table).Interior.ColorIndex = 4
            
                ElseIf qrtlyChange < 0 Then
                    ws.Range("J" & new_table).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & new_table).Interior.ColorIndex = xlNone
                End If
            'End If
    
            'Greatest increase/decrease/total vol
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                increaseTicker = ticker_name
            End If
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                decreaseTicker = ticker_name
            End If
        
            If stock_vol > greatestVol Then
                greatestVol = stock_vol
                totalVolTicker = ticker_name
            End If
            
            new_table = new_table + 1
        
            If i < lastRow Then
                openPrice = ws.Cells(i + 1, 3).Value
                qrtlyChange = 0
                stock_vol = 0
            End If
            
            Else
                stock_vol = stock_vol + ws.Cells(i, 7).Value
        End If
    Next i
        'ws.Range("O2").Value = "Greatest % Increase"
        'ws.Range("O3").Value = "Greatest % Decrease"
        'ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P2").Value = increaseTicker
        ws.Range("P3").Value = decreaseTicker
        ws.Range("P4").Value = totalVolTicker
        
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = greatestVol
                                                                                
Next ws
        MsgBox "Mission complete!"
    
End Sub

