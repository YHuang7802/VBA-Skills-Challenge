Attribute VB_Name = "Module1"
Sub multiple_year_stock():

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
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

    ws.Range("A1").Value = "<ticker>"

'headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'column format
ws.Columns("i:q").EntireColumn.AutoFit
ws.Columns("k").NumberFormat = "0.00%"
ws.Columns("j").NumberFormat = "0.00"

'cell text format
ws.Range("i:q").HorizontalAlignment = xlCenter

'integer for loop
stock = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'initialize the vol counter
    
vol = 0
    
'for loop
For i = 2 To lastrow
    
    Ticker = ws.Cells(i, 1).Value
    
    'i+1 to distinguish from previous ticker
    
    anoticker = ws.Cells(i + 1, 1).Value
    
    'conditional for volume summation
    
    If vol = 0 Then
    
        yr_open = ws.Cells(i, 3).Value
    
    End If
    
    'summing volume from previous cell
    vol = vol + ws.Cells(i, 7).Value
   
    'conditional for ticker and total volume sortation
    
    If Ticker <> anoticker Then

        ws.Cells(stock, 9).Value = Ticker
        ws.Cells(stock, 12).Value = vol
        
        yr_close = ws.Cells(i, 6).Value
        
    'percent change & yearly change, use closing and opening value
    
        ws.Cells(stock, 10).Value = yr_close - yr_open
        ws.Cells(stock, 11).Value = (yr_close - yr_open) / yr_open
        
         stock = stock + 1

          vol = 0
          
    End If
    
Next i
    
    
    'color format the gains and losses
    'new loop for color format
For i = 2 To lastrow
   
    If ws.Cells(i, 10).Value > 0 Then

        ws.Cells(i, 10).Interior.ColorIndex = 4

    ElseIf ws.Cells(i, 10).Value < 0 Then

        ws.Cells(i, 10).Interior.ColorIndex = 3
    
     Else: ws.Cells(i, 10).Interior.ColorIndex = 0
        
    End If
    
Next i
    
'max value for bonus
    
ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("k:k"))
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("k:k"))
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("l:l"))
    
'loop for tickers corresponding to the values of column q
For i = 2 To lastrow
    
    newticker = ws.Cells(i, 9).Value
    
    If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
        ws.Cells(2, 16).Value = newticker
        
    End If
        
    If ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
        ws.Cells(3, 16).Value = newticker
        
    End If
        
    If ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
        ws.Cells(4, 16).Value = newticker
        
    End If
    
Next i
    
'formatting
ws.Range("q2:q3").NumberFormat = "0.00%"
    
Next ws

End Sub








