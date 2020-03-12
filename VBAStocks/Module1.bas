Attribute VB_Name = "Module1"
Sub alphabetical_testing()

    '|-- Declare the variables --|

    'Declare a variable to store the value of the last non-blank row
    Dim last_row As Long

    'Declare a variable to store the value of the last non-blank column
    Dim last_column As Long

    'Assign variable: value of last non-blank row as "last_row"
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    'Check that the last_row value was calculated correctly
    'MsgBox (last_row)
    
    
    'Assign variable: value of last non-blank column as "last_column"
    last_column = Cells(1, Columns.Count).End(xlToLeft).Column
    'Check that the last_column value was calculated correctly
    'MsgBox (last_column)



End Sub
