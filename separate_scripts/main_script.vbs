
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
                
                'If the yearly stock price started at 0, set percent_change equal to "Error"
                Else
                    percent_change = "Error"
                    
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


End Sub