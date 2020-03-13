Sub alphabetical_testing()

    '|-- Declare the variables --|

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
    Dim percent_change As Double
    
    'Stores the total stock volume value for the current stock
    Dim stock_volume As Double
    
    '|-- Assign the variables --|

    'Set the summary index to start in row 2
    summary_index = 2
    
    'Set the stock volume variable to start at 0
    stock_volume = 0
    

    'Assign variable: value of last non-blank row as "last_row"
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    'Check that the last_row value was calculated correctly
    'MsgBox (last_row)
    
    'Assign variable: value of last non-blank column as "last_column"
    last_column = Cells(1, Columns.Count).End(xlToLeft).Column
    'Check that the last_column value was calculated correctly
    'MsgBox (last_column)

    'Add headers to the summary table
    Range("I1").Value = "First Row"
    Range("J1").Value = "Last Row"
    
    Range("K1").Value = "Ticker"
    Range("L1").Value = "Yearly Change"
    Range("m1").Value = "Percent Change"
    Range("n1").Value = "Total Stock Volume"

    'Loop through all of the ticker symbols in column 1
    For i = 2 To last_row:
        
        'Set a variable to track the current row number
        
        'Check if the next cell's ticker symbol is different than the current cell's ticker symbol
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'If so, assign the current ticker value to the 'ticker_symbol' variable
            Range("K" & summary_index).Value = Cells(i, 1).Value
            
            'Determine the last row for this ticker
            ticker_last = i
            
            'Add the stock volume of the current row to the stock_volume variable
            stock_volume = stock_volume + Range("G" & i).Value
            
            'Add the total stock volume value to the summary table in column "N"
            Range("N" & summary_index).Value = stock_volume
            
            'Add the yearly change value to the summary table
            Range("L" & summary_index).Value = yearly_change
            
            'Add the percent change value to the summary table
            Range("M" & summary_index).Value = percent_change
            
            'Add the ticker_last value to the left of the summary table to verify
            Range("J" & summary_index).Value = ticker_last
            
            'Calculate the yearly change value
            yearly_change = (Range("F" & ticker_last) - Range("C" & ticker_first).Value)
            
            'Add the yearly change value to the summary table in column "L"
            Range("l" & summary_index).Value = yearly_change
            
            'Check if the yearly change was positive
            If yearly_change > 0 Then
                
                'If so, set the cell background color to green
                Range("l" & summary_index).Interior.ColorIndex = 4
            
            'Check if the yearly change was negative
            ElseIf yearly_change < 0 Then
                
                'If so, set the cell background color to red
                Range("l" & summary_index).Interior.ColorIndex = 3
            
            'Check if the yearly change was zero
            ElseIf yearly_change = 0 Then
                
                'If so, set the cell background color to Blue
                Range("l" & summary_index).Interior.ColorIndex = 5
            
            End If
            
            'Calculate the percent change
            percent_change = (yearly_change / Range("C" & ticker_first).Value)
            
            'Add the percent change value to the summary_table in column "M"
            Range("M" & summary_index).Value = percent_change
            
            
            'Reset the stock volume value
            stock_volume = 0
            
            'Add 1 to the summary index
            summary_index = summary_index + 1
            
        'Check if this is the first row of a new ticker
        ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
            'If so, assign the current row to the ticker_first variable
            ticker_first = i
            
            'Add the ticker_first value to the left of the summary table to verify
            Range("I" & summary_index).Value = ticker_first
              
            'Add the stock volume of the current row to the stock_volume variable
            stock_volume = stock_volume + Range("G" & i).Value
            
            
        Else
            'Add the stock volume of the current row to the stock_volume variable
            stock_volume = stock_volume + Range("G" & i).Value
        
        End If
        
    'Check the next row
    Next i
    
    'Loop through the yearly change and percent change values in the summary table
    For i = 2 To Range("L" & Rows.Count).End(xlUp).Row
    
        'If the value in each cell is greater than 0 color the cell background green
    
        'If the value in the cell is
    
    Next i



End Sub

