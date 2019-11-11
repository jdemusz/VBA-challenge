Sub StockTrade()

    Dim Ticker_Symbol As String
    Dim Stock_Total As Double
    Dim Summary_Table_Row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim Yearly_change As Double
    Dim percentage_change As Double
    Dim Lastrow As Long

    For Each ws In Worksheets
        If ws.Name <> "A" Then
            Exit For
        End If

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Stock_Total = 0
        Summary_Table_Row = 2
        Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        open_price = ws.Cells(2, 3).Value

        For i = 2 To Lastrow

            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker Symblo
                Ticker_Symbol = ws.Cells(i, 1).Value
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                'Set open & close price
                close_price = ws.Cells(i, 6).Value
                Yearly_change = close_price - open_price

                If open_price <> 0 Then
                    percentage_change = Yearly_change / open_price
                Else
                    percentage_change = 0
                End If

                ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
                ws.Range("J" & Summary_Table_Row).Value = Yearly_change
                ws.Range("K" & Summary_Table_Row).Value = percentage_change
                ws.Range("L" & Summary_Table_Row).Value = Stock_Total

                If Yearly_change > 0 Then
                    ' Set the Font color to green
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf Yearly_change < 0 Then
                    ' Set the Font color to red
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If

                'Format % for Yearly_change
                Cells(i, 11).Value = Format(Yearly_change, "Percent")

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                ' Reset the Stock Total
                Stock_Total = 0
                open_price = ws.Cells(i + 1, 3).Value
                ' If the cell immediately following a row is the same...
            Else
                ' Add to the Stock Total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            End If

        Next i

    Next ws

End Sub
