Sub ticker():

For Each ws In Worksheets

Dim ticker_sym As String
Dim vol_total As Double
Dim table_row As Double
Dim first_open As Double
Dim last_close As Double
Dim change As Double
Dim perc_change As Double
Dim perc_zero As Double


'create headers for the new columns

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

    
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

table_row = 2

first_open = ws.Cells(2, 3).Value

For I = 2 To lastrow

    If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then

       vol_total = vol_total + ws.Cells(I, 7).Value
        ticker_sym = ws.Cells(I, 1).Value
        ws.Cells(table_row, 9).Value = ticker_sym
        ws.Cells(table_row, 12).Value = vol_total
    
        ' Cells(i,6).value should be the value for closing price at end of year for the stock

        last_close = ws.Cells(I, 6).Value
    
        'must use the first_open now before it moves to the next ticker symbol
    
        'calculate and display the amount changed
    
        change = last_close - first_open
        ws.Cells(table_row, 10).Value = change
    
        'add conditional formatting
    
        If change > 0 Then
            ws.Cells(table_row, 10).Interior.ColorIndex = 4
    
        ElseIf change < 0 Then
            ws.Cells(table_row, 10).Interior.ColorIndex = 3
        
        End If
        
        'calculate and display the percentage changed
    
        'need to fix if first_open is zero
    
        perc_zero = Empty
    
        If first_open = 0 Then
            ws.Cells(table_row, 11).Value = perc_zero
            ws.Cells(table_row, 11).HorizontalAlignment = xlRight
        
        Else
            perc_change = (change / first_open)
            ws.Cells(table_row, 11).NumberFormat = "0.00%"
            ws.Cells(table_row, 11).Value = perc_change
        
        End If
        
        'update the first open price
        
        first_open = ws.Cells(I + 1, 3).Value
        ws.Cells(table_row, 12).Value = vol_total
    
        'clear the total and ticker symbol for the next set
        
        vol_total = 0
        ticker_sym = ""
        
        table_row = table_row + 1

    Else
        vol_total = vol_total + ws.Cells(I, 7).Value
        
    End If

Next I
    
'BONUS SECTION
 
'create headers for the bonus items
    
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

Dim great_inc As Double
Dim great_inc_tick As String
Dim great_dec As Double
Dim great_dec_tick As String
Dim great_vol As Double
Dim great_vol_tick As String

lastrow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row

great_inc = ws.Cells(2, 11)
great_inc_tick = ws.Cells(2, 9)
great_dec = ws.Cells(2, 11)
great_dec_tick = ws.Cells(2, 9)
great_vol = ws.Cells(2, 12)
great_vol_tick = ws.Cells(2, 9)

For x = 2 To lastrow2

    If great_inc < ws.Cells(x, 11).Value Then
        great_inc = ws.Cells(x, 11)
        great_inc_tick = ws.Cells(x, 9)
    End If
    
    If great_dec > ws.Cells(x, 11).Value Then
        great_dec = ws.Cells(x, 11)
        great_dec_tick = ws.Cells(x, 9)
    End If
    
    If great_vol < ws.Cells(x, 12).Value Then
        great_vol = ws.Cells(x, 12)
        great_vol_tick = ws.Cells(x, 9)
    End If
    
Next x

ws.Cells(2, 16).Value = great_inc_tick
ws.Cells(2, 17).Value = great_inc
ws.Cells(2, 17).NumberFormat = "0.00%"

ws.Cells(3, 16).Value = great_dec_tick
ws.Cells(3, 17).Value = great_dec
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Cells(4, 16).Value = great_vol_tick
ws.Cells(4, 17).Value = great_vol

Next ws

End Sub
