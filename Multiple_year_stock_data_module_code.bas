Attribute VB_Name = "Module1"
Sub stonks_go_up():

'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.

'define variables
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Long
Dim Total_Stock_Volume As Long
Dim vol As Double
Dim yr_open As Double
Dim yr_close As Double
Dim time As Double
Dim lastrow As Double

'headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'column format
Columns("i:q").EntireColumn.AutoFit
Columns("k").NumberFormat = "0.00%"
Columns("j").NumberFormat = "0.00"

'cell text format
Range("i:q").HorizontalAlignment = xlCenter

'integer for loop
stock = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'initialize the vol counter
    
vol = 0
    
'for loop
For i = 2 To lastrow
    
    Ticker = Cells(i, 1).Value
    
    'i+1 to distinguish from previous ticker
    
    anoticker = Cells(i + 1, 1).Value
    
    'conditional for volume summation
    
    If vol = 0 Then
    
        yr_open = Cells(i, 3).Value
    
    End If
    
    'summing volume from previous cell
    vol = vol + Cells(i, 7).Value
   
    'conditional for ticker and total volume sortation
    
    If Ticker <> anoticker Then

        Cells(stock, 9).Value = Ticker
        Cells(stock, 12).Value = vol
        
        yr_close = Cells(i, 6).Value
        
    'percent change & yearly change, use closing and opening value
    
        Cells(stock, 10).Value = yr_close - yr_open
        Cells(stock, 11).Value = (yr_close - yr_open) / yr_open
        
         stock = stock + 1

          vol = 0
          
    End If
    
Next i
    
    
    'color format the gains and losses
    'new loop for color format
For i = 2 To lastrow
   
    If Cells(i, 10).Value > 0 Then

        Cells(i, 10).Interior.ColorIndex = 4

    ElseIf Cells(i, 10).Value < 0 Then

        Cells(i, 10).Interior.ColorIndex = 3
    
     Else: Cells(i, 10).Interior.ColorIndex = 0
        
    End If
    
Next i
    
'max value for bonus
    
Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("k:k"))
Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("k:k"))
Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("l:l"))
    
'loop for tickers corresponding to the values of column q
For i = 2 To lastrow
    
    newticker = Cells(i, 9).Value
    
    If Cells(i, 11).Value = Cells(2, 17).Value Then
        Cells(2, 16).Value = newticker
        
    End If
        
    If Cells(i, 11).Value = Cells(3, 17).Value Then
        Cells(3, 16).Value = newticker
        
    End If
        
    If Cells(i, 12).Value = Cells(4, 17).Value Then
        Cells(4, 16).Value = newticker
        
    End If
    
Next i
    
'formatting
Range("q2:q3").NumberFormat = "0.00%"
    
End Sub




