Sub ticker():

Dim ticker_sym As String
Dim vol_total As Double
Dim table_row As Double
Dim first_open As Double
Dim last_close As Double
Dim change As Double
Dim perc_change As Double
Dim perc_zero As Double


'create headers for the new columns

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

    
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'lastrow = 500

table_row = 2

first_open = Cells(2, 3).Value

For i = 2 To lastrow

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

       vol_total = vol_total + Cells(i, 7).Value
        ticker_sym = Cells(i, 1).Value
        Cells(table_row, 9).Value = ticker_sym
        Cells(table_row, 12).Value = vol_total
    
        ' Cells(i,6).value should be the value for closing price at end of year for the stock
        last_close = Cells(i, 6).Value
    
        'must use the first_open now before it moves to the next ticker symbol
    
        'calculate and display the amount changed
    
        change = last_close - first_open
        Cells(table_row, 10).Value = change
    
        'add conditional formatting
    
        If change > 0 Then
            Cells(table_row, 10).Interior.ColorIndex = 4
    
        ElseIf change < 0 Then
            Cells(table_row, 10).Interior.ColorIndex = 3
        
        End If
        
        'calculate and display the percentage changed
    
        'need to fix if first_open is zero
    
        perc_zero = Empty
    
        If first_open = 0 Then
            Cells(table_row, 11).Value = perc_zero
            Cells(table_row, 11).HorizontalAlignment = xlRight
        
        Else
            perc_change = (change / first_open)
            Cells(table_row, 11).NumberFormat = "0.00%"
            Cells(table_row, 11).Value = perc_change
        
        End If
        
        'update the first open price
        
        first_open = Cells(i + 1, 3).Value
        Cells(table_row, 12).Value = vol_total
    
        'clear the total and ticker symbol for the next set
        
        vol_total = 0
        ticker_sym = ""
        
        table_row = table_row + 1

    Else
        vol_total = vol_total + Cells(i, 7).Value
        
    End If

Next i
    
    
'BONUS SECTION
 
'create headers for the bonus items
    
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Dim great_inc As Double
Dim great_inc_tick As String
Dim great_dec As Double
Dim great_dec_tick As String
Dim great_vol As Double
Dim great_vol_tick As String

lastrow2 = Cells(Rows.Count, 11).End(xlUp).Row

great_inc = Cells(2, 11)
great_inc_tick = Cells(2, 9)
great_dec = Cells(2, 11)
great_dec_tick = Cells(2, 9)
great_vol = Cells(2, 12)
great_vol_tick = Cells(2, 9)


For x = 2 To lastrow2

    If great_inc < Cells(x, 11).Value Then
        great_inc = Cells(x, 11)
        great_inc_tick = Cells(x, 9)
    End If
    
    If great_dec > Cells(x, 11).Value Then
        great_dec = Cells(x, 11)
        great_dec_tick = Cells(x, 9)
    End If
    
    If great_vol < Cells(x, 12).Value Then
        great_vol = Cells(x, 12)
        great_vol_tick = Cells(x, 9)
    End If
    
Next x

Cells(2, 16).Value = great_inc_tick
Cells(2, 17).Value = great_inc
Cells(2, 17).NumberFormat = "0.00%"

Cells(3, 16).Value = great_dec_tick
Cells(3, 17).Value = great_dec
Cells(3, 17).NumberFormat = "0.00%"

Cells(4, 16).Value = great_vol_tick
Cells(4, 17).Value = great_vol

End Sub
