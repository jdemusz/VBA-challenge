Sub StockTrade()

'Part 1: Ticker Symbol & Vol.
'Part 2: Yearly_change calculation and format

    For Each ws In Worksheets
    
      ' Set an initial variable for holding the ticker
      Dim Ticker_Symbol As String
    
      ' Set an initial variable for holding the total vol. per stock
      Dim Stock_Total As LongLong
      Stock_Total = 0
    
      ' Keep track of the location for each stock in the summary table
      Dim Summary_Table_Row As Long
      Summary_Table_Row = 2
      
      ' Set date as string
      Dim open_price As Double
      Dim close_price As Double
      Dim Yearly_change As Double
          
      ' Loop through all stock trade
      Dim Lastrow As Long
      Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      For i = 2 To Lastrow
    
        ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Set the ticker Symblo
          Ticker_Symbol = ws.Cells(i, 1).Value
    
          ' Add to the Stock Total
          Stock_Total = Ticker_Total + ws.Cells(i, 7).Value
    
          ' Print the Ticker Symbol in the Summary Table
          ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
    
          ' Print the Stock Total to the Summary Table
          ws.Range("L" & Summary_Table_Row).Value = Stock_Total
    
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Stock Total
          Stock_Total = 0
    
        ' If the cell immediately following a row is the same brand...
        Else
    
         ' Add to the Stock Total
          Stock_Total = Stock_Total + ws.Cells(i, 7).Value
    
        End If
    
      Next i
    
    Next ws

End Sub
