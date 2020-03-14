Sub main_script()
    
    'Loop through each worksheet
    For Each ws In Worksheets
        'Declare variables
        'Stores the last non-blank row number
        Dim last_row As Long

        'Stores the last non-blank column number
        Dim last_column As Long
        
        'Stores the row index for the summary table
        Dim summary_index As Long

        'Stores the ticker symbol of the current row
        Dim ticker_symbol As String

        'Stores the yearly change value for the current stock
        Dim yearly_change As Double

        'Stores the percent change value for the current stock
        Dim percent_change As Variant
        
        'Stores the total stock volume value for the current stock
        Dim stock_volume As Double
        
        'Assign variables

        'Set the summary index to start in row 2
        summary_index = 2
        
        'Set the stock volume variable to start at 0
        stock_volume = 0

        'Assign variable: value of last non-blank row as "last_row"
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Check that the last_row value was calculated correctly
        'MsgBox (last_row)
        
        'Assign variable: value of last non-blank column as "last_column"
        last_column = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        'Check that the last_column value was calculated correctly
        'MsgBox (last_column)

        'Add headers to the summary table
        ws.Range("I1").Value = "First Row"
        ws.Range("J1").Value = "Last Row"
        
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("m1").Value = "Percent Change"
        ws.Range("n1").Value = "Total Stock Volume"

        'Loop through all of the ticker symbols in column 1
        For i = 2 To last_row:
            
            'Set a variable to track the current row number
            
            'Check if the next cell's ticker symbol is different than the current cell's ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'If so, assign the current ticker value to the 'ticker_symbol' variable
                ws.Range("K" & summary_index).Value = Cells(i, 1).Value
                
                'Determine the last row for this ticker
                ticker_last = i
                
                'Add the stock volume of the current row to the stock_volume variable
                stock_volume = stock_volume + ws.Range("G" & i).Value
                
                'Add the total stock volume value to the summary table in column "N"
                ws.Range("N" & summary_index).Value = stock_volume
                
                'Add the yearly change value to the summary table
                ws.Range("L" & summary_index).Value = yearly_change
                
                'Add the percent change value to the summary table
                ws.Range("M" & summary_index).Value = percent_change
                
                'Add the ticker_last value to the left of the summary table to verify
                ws.Range("J" & summary_index).Value = ticker_last
                
                'Calculate the yearly change value
                yearly_change = ws.Range("F" & ticker_last).Value - ws.Range("C" & ticker_first).Value
                
                'Add the yearly change value to the summary table in column "L"
                ws.Range("l" & summary_index).Value = yearly_change
                
                'Check if the yearly change was positive
                If yearly_change > 0 Then
                    
                    'If so, set the cell background color to green
                    ws.Range("l" & summary_index).Interior.ColorIndex = 4
                
                'Check if the yearly change was negative
                ElseIf yearly_change < 0 Then
                    
                    'If so, set the cell background color to red
                    ws.Range("l" & summary_index).Interior.ColorIndex = 3
                
                'Check if the yearly change was zero
                ElseIf yearly_change = 0 Then
                    
                    'If so, set the cell background color to Blue
                    ws.Range("l" & summary_index).Interior.ColorIndex = 5
                
                End If
                
                'Check that the yearly opening price of the current stock is not zero
                If ws.Range("C" & ticker_first).Value <> 0 Then
                
                    'Calculate the percent change
                    percent_change = yearly_change / ws.Range("C" & ticker_first).Value
                
                'If the yearly stock price started at 0, set percent_change equal to "0"
                Else
                    percent_change = 0
                    
                End If
                
                'Add the percent change value to the summary_table in column "M"
                ws.Range("M" & summary_index).Value = percent_change
                
                'Format column "M" as percent
                ws.Range("M" & summary_index).NumberFormat = "0.00%"
                
                'Reset the stock volume value
                stock_volume = 0
                
                'Add 1 to the summary index
                summary_index = summary_index + 1
                
            'Check if this is the first row of a new ticker
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'If so, assign the current row to the ticker_first variable
                ticker_first = i
                
                'Add the ticker_first value to the left of the summary table to verify
                ws.Range("I" & summary_index).Value = ticker_first
                
                'Add the stock volume of the current row to the stock_volume variable
                stock_volume = stock_volume + ws.Range("G" & i).Value
                
                
            Else
                'Add the stock volume of the current row to the stock_volume variable
                stock_volume = stock_volume + ws.Range("G" & i).Value
            
            End If
            
        'Check the next row
        Next i
    
    'Move on to the next worksheet
    Next ws

'The worksheets were combined here
'---------------------------------
    
    'Loop through each worksheet
    For Each ws In Worksheets
        
        'Set up the headers for the second summary table
        ws.Range("Q2").Value = "Greatest Increase"
        ws.Range("Q3").Value = "Greatest Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"
        ws.Range("R1").Value = "Ticker"
        ws.Range("S1").Value = "Value"

        'Reset the counter variables
        max_increase = 0
        max_decrease = 0
        greatest_volume = 0
        increase_ticker = ""
        decrease_ticker = ""
        volume_ticker = ""

        
        'Loop through the percent change column in the summary table
        For i = 2 To ws.Range("m" & Rows.Count).End(xlUp).Row
            
            'Check if the current cell value is greater than max_increase
            If ws.Range("m" & i).Value > max_increase Then
                
                'Set max increase equal to the current cell value
                max_increase = ws.Range("m" & i).Value
                'Set increase_ticker equal to the ticker value for the current row
                increase_ticker = ws.Range("k" & i).Value
            
            'Check if the current cell value is less than max_decrease
            ElseIf ws.Range("m" & i).Value < max_decrease Then
            
                'Set max_decrease equal to the current cell value
                max_decrease = ws.Range("m" & i).Value
                'Set decrease_ticker equal to the ticker value for the current row
                decrease_ticker = ws.Range("k" & i).Value

            End If
            
        Next i
        
        'Loop through the total stock volume column in the summary table
        For i = 2 To ws.Range("n" & Rows.Count).End(xlUp).Row
        
            'Check if the current cell is greater than greatest_volume
            If ws.Range("n" & i).Value > greatest_volume Then
                
                'Set greatest_volume equal to the current cell value
                greatest_volume = ws.Range("n" & i).Value
                'Set volume_ticker equal to the ticker value for the current row
                volume_ticker = ws.Range("k" & i).Value
            
            End If
            
        Next i
        
        'Add the value of the greatest percent increase to summary table 2
        ws.Range("s2").Value = max_increase
        'Add the ticker symbol of the stock with the greatest percent increase to summary table 2
        ws.Range("r2").Value = increase_ticker
        
        'Add the value of the greatest percent decrease to summary table 2
        ws.Range("s3").Value = max_decrease
        'Add the ticker symbol of the stock with the greatest percent decrease to summary table 2
        ws.Range("r3").Value = decrease_ticker
        
        'Add the value of the greatest total stock volume to summary table 2
        ws.Range("s4").Value = greatest_volume
        'Add the ticker symbol of the stock with the greatest volume to summary table 2
        ws.Range("r4").Value = volume_ticker
        
        'Format the greatest percent increase and decrease cells in summary table 2 as percent
        ws.Range("S2:S3").NumberFormat = "0.00%"
       
    
    Next ws


End Sub