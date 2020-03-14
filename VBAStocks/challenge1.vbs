Sub testing()
    
    'Loop through each worksheet
    for each ws in worksheets
        
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
            If ws.Range("m" & i).Value > max_increase and ws.Range("m" & i).Value <> "Error" Then
                
                'Set max increase equal to the current cell value
                max_increase = ws.Range("m" & i).Value
                'Set increase_ticker equal to the ticker value for the current row
                increase_ticker = ws.Range("k" & i).Value
            
            'Check if the current cell value is less than max_decrease
            ElseIf ws.Range("m" & i).Value < max_decrease and ws.Range("m" & i).Value <> "Error" Then
            
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
    
    next ws

End Sub