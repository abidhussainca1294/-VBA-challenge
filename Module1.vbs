Attribute VB_Name = "Module1"
 Sub vbchallenge()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' --------------------------------------------
        ' Extract distinct Ticker
        ' --------------------------------------------

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Creating an initial variable for holding the Ticker
         Dim Ticker As String
        
        'Creating initial variable for open price, close price, yearly chnage, percent change
        'Open_price_next stores the value for the next ticker
         Dim open_price As Double
         Dim close_price As Double
         Dim open_price_next As Double
         Dim yearly_change As Double
         Dim percent_change As Double
         
  'Tracking Ticker symbol location for the Ticker column
        Dim Ticker_Table_Row As Integer
        Ticker_Table_Row = 2
  
  ' Set an initial variable for holding the total stock volume
        Dim Total_stock_volume As Double
        
  ' initialising the variables
        Total_stock_volume = 0
        open_price_next = ws.Cells(2, 3).Value
        close_price = 0
     
  'Add column headings
        ws.Range("I1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
  
  ' Loop through Ticker column
  For i = 2 To LastRow
      
         
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker = ws.Cells(i, 1).Value
      
      ' Print the Ticker in the Ticker column
      ws.Range("I" & Ticker_Table_Row).Value = Ticker
      
      'set closed price
      close_price = ws.Cells(i, 6).Value
       
      'set open price next for the next ticker
      open_price_next = ws.Cells(i + 1, 3).Value
      
      'Calculating yearly change
      yearly_change = close_price - open_price
      
      'Print yearly change
      ws.Range("j" & Ticker_Table_Row).Value = yearly_change
      
      'Formatting cell as per the change red or green
      If yearly_change <= 0 Then
        ws.Range("j" & Ticker_Table_Row).Interior.ColorIndex = 3
      Else
        ws.Range("j" & Ticker_Table_Row).Interior.ColorIndex = 4
      End If
      
      'calculating percent change
      percent_change = (yearly_change / open_price)
      
      'Formatting colum as Percentage
       ws.Range("k" & Ticker_Table_Row).NumberFormat = "0.00%"
       
       'Print percent change
       ws.Range("k" & Ticker_Table_Row).Value = percent_change
     
      ' Add to the Total_stock_volume
      Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value

      ' Print the Total_stock_volume  to the Total_stock_volume column
      ws.Range("l" & Ticker_Table_Row).Value = Total_stock_volume

      ' Add one to the Ticker table row
      Ticker_Table_Row = Ticker_Table_Row + 1
      
      ' Reset the  Total stock volume
      Total_stock_volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the  Total stock volume
       Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
       
       'updating open price with the open price value obtained for the next ticker at the starting of year
       open_price = open_price_next
    
    End If
      
Next i

       ' --------------------------------------------
        ' Adding functionality for Greatest % increase, Greatest % decrease and Greatest total volume
        ' --------------------------------------------

       'Creating initial variable for greatest increase , greatest decrease and associated ticker
        Dim min_change As Double
        Dim max_change As Double
        Dim ticker_max As String
        Dim ticker_min As String
       
       'Creating variables which stores the greatest increase and greatest decrease
        Dim max As Double
        Dim min As Double
        
       'Initialising the variable for determining greatest % increase and greatest % decrease
        min_change = 0
        max_change = 0
        
        'Determining the Last row for percentage change column
         LastRow_change = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Row and column header for the summary table
         ws.Range("P1").Value = "Ticker"
         ws.Range("Q1").Value = "value"
         ws.Range("O2").Value = "Greatest % Increase"
         ws.Range("O3").Value = "Greatest % Decrease"
         ws.Range("O4").Value = "Greatest Total Volume"

    ' Loop through Percentage change column
      For j = 2 To LastRow_change
  
    'Finding the greatest % increase
  
     'Check if the value in the cell is more than max_change
       If ws.Cells(j, 11).Value > max_change Then
     
     'Update max_change with the cell value and ticker_max with associated ticker
      max_change = ws.Cells(j, 11).Value
      ticker_max = ws.Cells(j, 9).Value
        
     End If
    
     'Update max with the greatest value in the max_change during the loop
      max = max_change
    
    'Finding the greatest % decrease
 
      'Check if the value in the cell is less than min_change
         If ws.Cells(j, 11).Value < min_change Then
         
      'Update min_change with the cell value and ticker_min with associated ticker
         min_change = ws.Cells(j, 11).Value
         ticker_min = ws.Cells(j, 9).Value
        
     End If
  
      'Update min with the greatest decrease value in the min_change during the loop
         min = min_change
    Next j
    
    'Formatting associated cell as Percentage
     ws.Range("Q2:Q3").NumberFormat = "0.00%"
    'Print results for Greatest% increase and Greatest percent decrease
     ws.Range("Q2") = max
     ws.Range("P2") = ticker_max
     ws.Range("Q3") = min
     ws.Range("P3") = ticker_min
  
  ' Alternative method for finding the maximum, Used for obtaining Greatest Total Volume in the Total stock Volume
  ' --------------------------------------------
    
    'Creating variable for maximum value and the row associated with maximum value found in the column
    Dim max_value As Double
    Dim max_row As Integer
    
    'Finding maximum in the Total Stock Volume
     max_value = Application.WorksheetFunction.max(ws.Columns("l"))
    
    'Finding corresponding row for maximum value
    max_row = Application.WorksheetFunction.Match(max_value, ws.Columns("l"), 0)
    
    'Print Greatest Total Volume value in the summary table
     ws.Range("Q4") = max_value
     
     'Print corrsponding ticker with the Greatest total Volume
    ws.Range("P4") = ws.Cells(max_row, 9).Value
  
  
Next ws
      

      
End Sub
 

