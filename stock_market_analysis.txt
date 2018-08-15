Sub stockAnalysis()
    'declare variables - rowNum(for printing unique value),first and last for earliest and latest date
    Dim totalVolume As Double
    Dim ticker As String
    Dim rowNum As Double
    Dim first As Double
    Dim last As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestTotal As Double
    
    'loop through each sheet
    For Each ws In Worksheets
        'initialize variables
        totalVolume = 0
        first = 0
        last = 0
        yearlyChange = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestTotal = 0
        rowNum = 2
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("I1").Value = "Ticker"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
                
        'loop through each year of stock data and grab the TotalVolume and TickerSymbol that coincide with it
        For i = 2 To lastRow
          
            If i = 2 Then
                first = ws.Cells(i, 3).Value
            End If
            
            If Trim(ws.Cells(i + 1, 1).Value) <> Trim(ws.Cells(i, 1).Value) Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ticker = ws.Cells(i, 1).Value
                last = ws.Cells(i, 6).Value
                ws.Cells(rowNum, 9).Value = ticker
                ws.Cells(rowNum, 12).Value = totalVolume
                ws.Cells(rowNum, 10) = last - first
                'To manage divide by zero error
                If first = 0 Then
                    percentChange = last
                Else
                   percentChange = ((last - first) / first)
                End If
                ws.Cells(rowNum, 11).Value = percentChange
                ws.Cells(rowNum, 11).NumberFormat = "0.00%"
                If ws.Cells(rowNum, 11).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Cells(rowNum, 11).Value
                    ws.Range("P2").Value = ws.Cells(rowNum, 9).Value
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                ElseIf ws.Cells(rowNum, 11).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Cells(rowNum, 11).Value
                    ws.Range("P3").Value = ws.Cells(rowNum, 9).Value
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                End If
                If ws.Cells(rowNum, 12).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Cells(rowNum, 12).Value
                    ws.Range("P4").Value = ws.Cells(rowNum, 9).Value
                End If
                If ws.Cells(rowNum, 10) > 0 Then
                    ws.Cells(rowNum, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowNum, 10).Interior.ColorIndex = 3
                End If
                first = ws.Cells(i + 1, 3).Value
                rowNum = rowNum + 1
                totalVolume = 0
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
        Next i
                        
            ws.Columns("A:V").AutoFit
          
    Next ws

End Sub


