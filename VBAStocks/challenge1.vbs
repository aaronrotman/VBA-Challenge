Sub testing()

    'Set up the headers for the second summary table
    Range("Q2").Value = "Greatest Increase"
    Range("Q3").Value = "Greatest Decrease"
    Range("Q4").Value = "Greatest Total Volume"
    Range("R1").Value = "Ticker"
    Range("S1").Value = "Value"
    
    'Loop through the percent change column in the summary table
    For i = 2 To Range("m" & Rows.Count).End(xlUp).Row
        
        'Check if the current cell value is greater than max_increase
        If Range("m" & i).Value > max_increase Then
            
            'Set max increase equal to the current cell value
            max_increase = Range("m" & i).Value
            'Set increase_ticker equal to the ticker value for the current row
            increase_ticker = Range("k" & i).Value
        
        'Check if the current cell value is less than max_decrease
        ElseIf Range("m" & i).Value < max_decrease Then
        
            'Set max_decrease equal to the current cell value
            max_decrease = Range("m" & i).Value
            'Set decrease_ticker equal to the ticker value for the current row
            decrease_ticker = Range("k" & i).Value

        End If
        
    Next i
    
    'Loop through the total stock volume column in the summary table
    For i = 2 To Range("n" & Rows.Count).End(xlUp).Row
    
        'Check if the current cell is greater than greatest_volume
        If Range("n" & i).Value > greatest_volume Then
            
            'Set greatest_volume equal to the current cell value
            greatest_volume = Range("n" & i).Value
            'Set volume_ticker equal to the ticker value for the current row
            volume_ticker = Range("k" & i).Value
        
        End If
        
    Next i
    
    '|-- Add values to summary table 2 --|
    
    'Add the value of the greatest percent increase to summary table 2
    Range("s2").Value = max_increase
    'Add the ticker symbol of the stock with the greatest percent increase to summary table 2
    Range("r2").Value = increase_ticker
    
    'Add the value of the greatest percent decrease to summary table 2
    Range("s3").Value = max_decrease
    'Add the ticker symbol of the stock with the greatest percent decrease to summary table 2
    Range("r3").Value = decrease_ticker
    
    'Add the value of the greatest total stock volume to summary table 2
    Range("s4").Value = greatest_volume
    'Add the ticker symbol of the stock with the greatest volume to summary table 2
    Range("r4").Value = volume_ticker

End Sub