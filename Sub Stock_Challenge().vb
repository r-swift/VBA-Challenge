Sub Stock_Challenge()
Dim ticker As String
Dim number_tickers As String
Dim lastRow As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double

For Each ws In Worksheets
    ws.Activate
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    For i = 2 To lastRow
        ticker = Cells(i, 1).Value
        If opening_price = 0 Then
        opening_price = Cells(i, 3).Value
        End If
        
    total_stock_volume = total_stock_volume + Cells(i, 7).Value
        If Cells(i + 1, 1).Value <> ticker Then
        number_tickers = number_tickers + 1
        Cells(number_tickers + 1, 9) = ticker
        closing_price = Cells(i, 6)
        yearly_change = closing_price - opening_price
        Cells(number_tickers + 1, 10).Value = yearly_change
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(mumber_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
        If opening_price = 0 Then
            percent_change = 0
        Else
            percent_change = (yearly_change / opening_price)
        End If
            
        Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
    
        opening_price = 0
        Cells(number_tickers + 1, 12).Value = total_stock_volume
        total_stock_volume = 0
        End If
    Next i
    
    Range("O2").Value = "Greatest Percent Increase"
    Range("O3").Value = "Greatest Percent Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
              
    Range("Q2").Formula = WorksheetFunction.Max(Range(Cells(2, "K"), Cells(lastRow, "K")))
    Range("Q2").Value = Range("Q2").Value
    
    Range("Q3").Formula = WorksheetFunction.Min(Range(Cells(2, "K"), Cells(lastRow, "K")))
    Range("Q3").Value = Range("Q3").Value
             
    Range("Q4").Formula = WorksheetFunction.Max(Range(Cells(2, "L"), Cells(lastRow, "L")))
    Range("Q4").Value = Range("Q4").Value
    
        For i = 2 To lastRow
            If Cells(i, 11).Value = Range("Q2").Value Then
               Range("P2").Value = Cells(i, 9).Value
            
            ElseIf Cells(i, 11).Value = Range("Q3").Value Then
               Range("P3").Value = Cells(i, 9).Value
               
            ElseIf Cells(i, 12).Value = Range("Q4").Value Then
               Range("P4").Value = Cells(i, 9).Value
            
            End If
            
        Next i
   
Next ws
           
End Sub
