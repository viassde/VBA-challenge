' Required to compute:
' a - The ticker symbol.
' b - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' c - The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' d - The total stock volume of the stock.

' Set an initial variable for holding the STOCK name
Dim Stock_Name As String
    
' Set variables for holding different metrics per stock
Dim Stock_YrOpen, Stock_YrClose, Stock_YrDelta, Stock_YrPercent, Stock_Total As Double
Dim MaxChange, MinChange, MaxTotalVol, tmpMx, tmpMn, tmpVol As Double

' Keep track of the location for each stock in the summary table
Dim Summary_Table_Row As Integer

Sub StockMkt_1()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
        ' Determine Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' init volume, table row
        Stock_Total = 0
        Summary_Table_Row = 2
        
        ' open price for 1st stock
        Stock_YrOpen = ws.Cells(Summary_Table_Row, 3).Value
    
        ' headers for outputs table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'ws.Cells(1, 13).Value = "Init Yr Price"
        'ws.Cells(1, 14).Value = "Final Yr Price"
          
        ' Loop through all stocks in sheet
        For I = 2 To LastRow
    
        ' Check if we are still within the same stock
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
   
                ' Set the stock name
                Stock_Name = ws.Cells(I, 1).Value
                ' stock final price
                Stock_YrClose = ws.Cells(I, 6).Value
                ' Add to Total Vol
                Stock_Total = Stock_Total + ws.Cells(I, 7).Value
                
                ' Print Stock name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Stock_Name
                ' Print Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Stock_Total
                ws.Range("J" & Summary_Table_Row).Value = Stock_YrClose - Stock_YrOpen
                           
                If Stock_YrOpen <> 0# Then   'dont divide by 0; when divided by zero % change will show as "empty"
                    ws.Range("K" & Summary_Table_Row).Value = (Stock_YrClose - Stock_YrOpen) / Stock_YrOpen
                End If
                                
                If Stock_YrClose < Stock_YrOpen Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3 'red
                Else
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4 'green
                End If
                        
                ws.Cells(Summary_Table_Row, 11).Style = "Percent"
                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                'ws.Range("M" & Summary_Table_Row).Value = Stock_YrOpen
                'ws.Range("N" & Summary_Table_Row).Value = Stock_YrClose
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
          
                ' Reset the stock Total
                Stock_Total = 0
                ' Open price for next stock
                Stock_YrOpen = ws.Cells(I + 1, 3).Value
    
            ' If the cell immediately following a row is the same stock...
            Else
                ' Add to the Total Volume
                Stock_Total = Stock_Total + ws.Cells(I, 7).Value
    
            End If
    
        Next I
    
        ' compute max price change
        tmpMx = WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row - 1).Value)
        row_tmpMx = 1 + WorksheetFunction.Match(tmpMx, ws.Range("K2:K" & (Summary_Table_Row - 1)).Value, 0)
        name_tmpMx = ws.Range("I" & row_tmpMx)
        
        ' compute min price change
        tmpMn = WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row - 1).Value)
        row_tmpMn = 1 + WorksheetFunction.Match(tmpMn, ws.Range("K2:K" & (Summary_Table_Row - 1)).Value, 0)
        name_tmpMn = ws.Range("I" & row_tmpMn)
        
        ' compute max Volume
        tmpVol = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row - 1).Value)
        row_tmpVol = 1 + WorksheetFunction.Match(tmpVol, ws.Range("L2:L" & (Summary_Table_Row - 1)).Value, 0)
        name_tmpVol = ws.Range("I" & row_tmpVol)
           
        ' just change names from tmp to final (this is not required)
        MaxChange = tmpMx
        name_MaxChange = name_tmpMx
        MinChange = tmpMn
        name_MinChange = name_tmpMn
        MaxTotalVol = tmpVol
        name_MaxVol = name_tmpVol
    
        ' write the 3 Summary Metrics at top of each worksheet
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(1, 17).Value = "Ticker"
         
        ws.Cells(2, 18).Value = MaxChange
        ws.Cells(2, 18).Style = "Percent"
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(2, 17).Value = name_MaxChange
        ws.Cells(2, 16).Value = "Greatest % increase"
        
        ws.Cells(3, 18).Value = MinChange
        ws.Cells(3, 18).Style = "Percent"
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = name_MinChange
        ws.Cells(3, 16).Value = "Greatest % decrease"
        
        ws.Cells(4, 18).Value = MaxTotalVol
        ws.Cells(4, 17).Value = name_MaxVol
        ws.Cells(4, 16).Value = "Greatest Total Volume"
    
    Next ws  ' next worksheet
    
End Sub

