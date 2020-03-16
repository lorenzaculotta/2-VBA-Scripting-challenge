Sub VBA_Challenge():

'INSTRUCTIONS
'Create a script that will loop through all the stocks for one year and output the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.

For Each ws In Worksheets
 
 '-----------------------------------------
 'PART 1
 '-----------------------------------------
 
 
    'define headers in summary table
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'define variables for summary table
    Dim ticker_symbol As String
    Dim yearly_change As Double
    Dim percent_change As Double
            
    Dim summary_table_row As Integer 'a variable to increment summary table row as we loop through ticker categories
    summary_table_row = 2
            
    Dim total_stock As Double 'a variable to increment total stock volume as we loop through category
    total_stock = 0
            
    'define variables for open price and close price
    Dim open_price As Double
    Dim close_price As Double
          
    'define number of tickers in the same category to keep track of row number within category
    Dim ticker_number As Double
    ticker_number = 0
            
                                                            
    'Loop through rows
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Defines LastRow of entire sheet
            
    For i = 2 To LastRow
               
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                           
            'extrapolate ticker symbol and print in summary table
            ticker_symbol = ws.Cells(i, 1).Value
            ws.Range("I" & summary_table_row).Value = ticker_symbol
                  
            'calculate open price and close price for current category
            open_price = ws.Cells(i - ticker_number, 3).Value
            close_price = ws.Cells(i, 6).Value
                    
            'calculate yearly change and print in summary table
            yearly_change = close_price - open_price                   
            ws.Range("J" & summary_table_row).Value = yearly_change
                           
                'Color-code Yearly Change Column
                If yearly_change >= 0 Then
                    ws.Range("J" & summary_table_row).Interior.Color = vbGreen
                Else
                    ws.Range("J" & summary_table_row).Interior.Color = vbRed
                End If
                                   
                           
                           
            'calculate percent change and print in summary table, with "percent" formatting. Avoid proble with 0 (example PLNT)
            If (yearly_change = 0) Or (open_price = 0) Then
                ws.Range("K" & summary_table_row).Value = "Null"
                        
            Else
                percent_change = yearly_change / open_price                           
                ws.Range("K" & summary_table_row).Value = percent_change
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            End If
                    
            'calculate and print total stock volume in summary table.
            total_stock = total_stock + ws.Cells(i, 7).Value                        
            ws.Range("L" & summary_table_row).Value = total_stock
                       
            'PREPARE FOR FOLLOWING TICKER CATEGORY
                'increment summary row
                summary_table_row = summary_table_row + 1                            
                    
                'reset total stock
                total_stock = 0
                        
                'reset ticker_number
                ticker_number = 0
             
        Else
            total_stock = total_stock + ws.Cells(i, 7).Value
            ticker_number = ticker_number + 1
                   
        End If
            
    Next i
           
        



    '-----------------------------------------
    'PART 2 - CHALLENGES
    '-----------------------------------------
    'INSTRUCTIONS
    'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
    


    'define headers in summary table
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % increase"
    ws.Cells(3, 14).Value = "Greatest % decrease"
    ws.Cells(4, 14).Value = "Greatest total volume"
                    
                
    'define variables for greatest % increase and decrease and greatest total volume
    Dim max_perc As Double
    Dim min_perc As Double
    Dim max_volume As Double
        
    'Retrieve Greatest % increase value and ticker
    max_perc = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 16).Value = max_perc
    ws.Cells(2, 16).NumberFormat = "0.00%"
            
        'associate ticker to greatest % increase
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value = max_perc Then
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
            End If
        Next i           
            
    'Retrieve Greatest % decrease value and ticker
    min_perc = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 16).Value = min_perc
    ws.Cells(3, 16).NumberFormat = "0.00%"
                
        'associate ticker to greatest % decrease
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value = min_perc Then
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
            End If
        Next i
        
    'Retrieve Greatest total volume value and ticker
    max_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 16).Value = max_volume
                
        'associate ticker to greatest total volume
        For i = 2 To LastRow
            If ws.Cells(i, 12).Value = max_volume Then
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
            End If
        Next i
        
        
    'Autofit sheet
    ws.Columns.AutoFit


Next ws

End Sub