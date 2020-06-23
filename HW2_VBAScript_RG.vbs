Attribute VB_Name = "Module1"
Sub Stock():

' Loop through all sheets

For Each ws In Worksheets

    Dim WorksheetName As String
    'Determine last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Define ticker
    Dim StockTicker As String
    
    ' Initialize variables and set initial values for holding the return, opening price and volume
    Dim Stock_Volume As Double
    Stock_Volume = 0
    Dim Stock_Return As Double
    Stock_Return = 0
    Dim Stock_Open As Double
    Stock_Open = ws.Cells(2, 3).Value
    Dim Stock_close As Double
    
    
    ' Create variables for the max values in summary table
    Dim Max_Volume As Double
    Dim Max_Increase As Double
    Dim Max_Decrease As Double
    
    
    ' Keep track of the location and make headers
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    ws.Cells(1, 9).Value = "Stock Ticker"
    ws.Cells(1, 10).Value = "Total Return of stock"
    ws.Cells(1, 11).Value = "Total % return of stock"
    ws.Cells(1, 12).Value = "Total Volume of stock"
    
    ' Headers for greatest values
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    
    'Loop through tickers from row 2 to the last row on the sheet
    For i = 2 To LastRow
        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          ' Set the Ticker
          StockTicker = ws.Cells(i, 1).Value
          ' Add the final row of volume to total
          Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
          ' Calculate the return from closing and opening prices
          Stock_Return = Stock_close - Stock_Open
              
            ' Create the Summary Table
            ' Print the Ticker
            ws.Range("I" & Summary_Table_Row).Value = StockTicker
            ' Print the Return
            ws.Range("J" & Summary_Table_Row).Value = Stock_Return
            ' Print the Return %
            ws.Range("K" & Summary_Table_Row).Value = ((Stock_Return + 0.00001) / (Stock_Open + 0.00001))
            ' Print the Volume
            ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
               
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
    
        'Reset variables and set the opening of next stock price
          Stock_Open = ws.Cells(i + 2, 3).Value
          Stock_Volume = 0
          Stock_Return = 0
      
      ' If the cell immediately following a row is the same ticker...
        Else
    
          ' Add to the total volume of current stock
          Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
          ' Closing price for the day
          Stock_close = ws.Cells(i, 6).Value
    
        End If
      
      Next i
      
    'Loop through to add conditional formatting and maximum values
      For i = 2 To LastRow
        ' Format the returns as percentages
        ws.Cells(i, 11).Style = "Percent"
        
        ' Conditional formatting for the returns
        If ws.Cells(i, 11).Value > 0.2 Then
        ws.Cells(i, 11).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 11).Value > 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 6
        ElseIf ws.Cells(i, 11).Value < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
        
        
        ' Max values
        ' Maximum volume
        Max_Volume = Application.WorksheetFunction.Max(ws.Range("L2:L5000"))
        ws.Cells(4, 16).Value = Max_Volume
        ' Pull ticker via match index
        ws.Cells(4, 15).Value = Application.WorksheetFunction.Index(ws.Range("I2:I5000"), Application.WorksheetFunction.Match(Max_Volume, ws.Range("L2:L5000"), 0), 1)
        
        
        ' Maximum % increase
        Max_Increase = Application.WorksheetFunction.Max(ws.Range("K2:K5000"))
        ws.Cells(2, 16).Value = Max_Increase
        ' Format cell as percentage
        ws.Cells(2, 16).Style = "Percent"
        ' Pull ticker
        ws.Cells(2, 15).Value = Application.WorksheetFunction.Index(ws.Range("I2:I5000"), Application.WorksheetFunction.Match(Max_Increase, ws.Range("K2:K5000"), 0), 1)
        
        ' Maximum % decrease
        Max_Decrease = Application.WorksheetFunction.Min(ws.Range("K2:K5000"))
        ws.Cells(3, 16).Value = Max_Decrease
        ' Format cell as percentage
        ws.Cells(3, 16).Style = "Percent"
        ' Pull ticker
        ws.Cells(3, 15).Value = Application.WorksheetFunction.Index(ws.Range("I2:I5000"), Application.WorksheetFunction.Match(Max_Decrease, ws.Range("K2:K5000"), 0), 1)
    
    
        'Reset the variables before moving to next sheet
        Max_Volume = 0
        Max_Decrease = 0
        Max_Increase = 0
    
        End If
                
        Next i
    
        'Autofit to display data
        ws.Columns("A:P").AutoFit
  
  'Move to next worksheet
  Next ws
    

End Sub

