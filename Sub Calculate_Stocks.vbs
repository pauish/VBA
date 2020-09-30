Sub Calculate_Stocks():

'Navigate WS
For ws = 1 To Worksheets.Count
Worksheets(ws).Activate


'Variables
Dim ticker As String
Dim tickerID As Integer
Dim last_row As Long
Dim opening_price, closing_price, stock_volume, yearly_change, percent_yearly_change, total_volume As Double


'Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent change"
Cells(1, 12).Value = "Total Stock Volume"

    last_row = Cells(Rows.Count, "A").End(xlUp).Row

'Init variables
   
    ticker = " "
    tickerID = 0
    opening_price = 0
    closing_price = 0
    total_volume = 0

    yearly_change = 0
    percent_yearly_change = 0

    
'Loop

    For i = 2 To last_row

        ticker = Cells(i, 1).Value
        
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        
        total_volume = total_volume + Cells(i, 7).Value
        
        
        If Cells(i + 1, 1).Value <> ticker Then
            tickerID = tickerID + 1
            
            Cells(tickerID + 1, 9) = ticker
            
            closing_price = Cells(i, 6)
            yearly_change = closing_price - opening_price
            
            Cells(tickerID + 1, 10).Value = yearly_change
            
            If yearly_change = 0 Then
            percent_change = 0
            
            Else
            
            percent_change = (yearly_change / opening_price)
          
        
            End If
        
            Cells(tickerID + 1, 11) = percent_change
            Cells(tickerID + 1, 11).NumberFormat = "0.00%"
            Cells(tickerID + 1, 12).Value = total_volume
            
        
            total_volume = 0
            opening_price = 0
        
        End If
        
        'Colors
        
        If yearly_change < 0 Then
        Cells(tickerID + 1, 10).Interior.ColorIndex = 3
        
        ElseIf yearly_change > 0 Then
        
        Cells(tickerID + 1, 10).Interior.ColorIndex = 10
    
        
        End If

    
    Next i
    
'Headers
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest total volume"


'Variables 
    
last_row = Cells(Rows.Count, "i").End(xlUp).Row

Dim greatest_increase, greatest_decrease, greatest_volume As Double
Dim ticker_increase, ticker_decrease, greatest_ticker As String

'Init variables 
greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0

ticker_increase = ""
ticker_decrease = ""
greatest_ticker = ""



'Greatest increase

    For j = 2 To last_row
    
    If greatest_increase = 0 Then
        greatest_increase = Cells(j, 11).Value
        ticker_increase = Cells(j, 9).Value
        
        End If
         
        If Cells(j, 11).Value > greatest_increase Then
        greatest_increase = Cells(j, 11).Value
        ticker_increase = Cells(j, 9).Value
        
        
        End If
        
        Next j
        
        Cells(2, 17).Value = greatest_increase
         Cells(2, 17).NumberFormat = "0.00%"
         Cells(2, 16).Value = ticker_increase
        
        
'Greatest decrease

    For j = 2 To last_row
    
    If greatest_decrease = 0 Then
        greatest_decrease = Cells(j, 11).Value
        ticker_decrease = Cells(j, 9).Value
        
        End If
         
        If Cells(j, 11).Value < greatest_decrease Then
        greatest_decrease = Cells(j, 11).Value
        ticker_decrease = Cells(j, 9).Value
        
        
        End If
        
        Next j
        
        Cells(3, 17).Value = greatest_decrease
         Cells(3, 17).NumberFormat = "0.00%"
         Cells(3, 16).Value = ticker_decrease
        
        
'Greatest total volume

    For j = 2 To last_row
    
    If greatest_volume = 0 Then
        greatest_volume = Cells(j, 12).Value
        greatest_ticker = Cells(j, 9).Value
        
        End If
         
        If Cells(j, 12).Value > greatest_volume Then
        greatest_volume = Cells(j, 12).Value
        greatest_ticker = Cells(j, 9).Value
        
        
        End If
        
        Next j
        
        Cells(4, 17).Value = greatest_volume
        Cells(4, 16).Value = greatest_ticker
        
	Next ws

End Sub
    
    



