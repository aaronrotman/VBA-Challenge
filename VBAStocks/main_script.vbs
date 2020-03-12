'This is the VBA script for the VBA Challenge assignment 

Sub alphabetical_testing()

    '|-- Declare the variables --|

    'Declare a variable to store the value of the last non-blank row
    dim last_row as integer

    'Declare a variable to store the value of the last non-blank column
    dim last_column as integer

    'Store the last non-blank row value as an integer
    last_row = cells(rows.count, 1).end(xltoleft).row

    'Store the last non-blank column value as an integer
    last_column = cells(1, columns.count).end(xlup).column














End Sub