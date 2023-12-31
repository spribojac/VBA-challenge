Attribute VB_Name = "Module5"
Sub stocktracker()

    'declare data variables
    Dim tickername As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalstockvolume As LongLong
    Dim lastrow As Long
    Dim ws As Worksheet
    Dim summarytablerow As Integer
    Dim percentincrease As Double
    Dim percentdecrease As Double
    Dim stockvol As LongLong
    
    'loop through all worksheets
    For Each ws In Worksheets
    
        'determine the lastrow
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
        
        'create new columns for ticker name, yearly change, percent change, total stock volume
        ws.Range("I1").EntireColumn.Insert
        ws.Cells(1, 9).Value = "Ticker Name"
        ws.Range("I1").EntireColumn.AutoFit
        
        ws.Range("J1").EntireColumn.Insert
        ws.Cells(1, 10).Value = "Yearly Change($)"
        ws.Range("J1").EntireColumn.AutoFit
        
        ws.Range("K1").EntireColumn.Insert
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Range("K1").EntireColumn.AutoFit
        
        ws.Range("L1").EntireColumn.Insert
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Range("L1").EntireColumn.AutoFit
        
        'Create second summary table for values
        ws.Range("O1").EntireColumn.Insert
        ws.Cells(1, 15).Value = "Ticker"
        ws.Range("O1").EntireColumn.AutoFit
        
        ws.Range("P1").EntireColumn.Insert
        ws.Cells(1, 16).Value = "Value"
        ws.Range("P1").EntireColumn.AutoFit
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'allow each ticker a new location in the summary table
        summarytablerow = 2
        
        'set values to hold yearly change, percent change, total stock volume, open value and close value
        yearlychange = 0
        percentchange = 0
        totalstockvolume = 0
        firstrow = 2
        percentincrease = 0
        percentdecrease = 0
        stockvol = 0
        ticker1 = 0
        ticker2 = 0
        ticker3 = 0
        
        'go through each ticker input
        For i = 2 To lastrow
        
            'check if the current row and following row are the same ticker name. If not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Create new ticker name and set variable
                tickername = ws.Cells(i, 1).Value
                
                'Put new ticker name in the summary table
                ws.Range("I" & summarytablerow).Value = tickername
                
                'Add to the total stock volume
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
                
                'Put new stock volume in summary table
                ws.Range("L" & summarytablerow).Value = totalstockvolume
                
                'reset the total stock volume
                totalstockvolume = 0
                    
                'Obtain the yearly change by subtracting the last close value from the first open value
                yearlychange = ws.Cells(i, 6).Value - ws.Cells(firstrow, 3).Value
                
                'Put the Yearly Change for previous row in summary table
                ws.Range("J" & summarytablerow).Value = yearlychange
                
                'Calculate the percentage change
                percentchange = (ws.Cells(i, 6).Value - ws.Cells(firstrow, 3).Value) / ws.Cells(firstrow, 3).Value
            
                'Put new percent change in the summary table
                ws.Range("K" & summarytablerow).Value = FormatPercent(percentchange)
                
                'add one to summary table row for the next new ticker name so they do not overwrite
                summarytablerow = summarytablerow + 1
            
                'update firstrow counter to move to next row
                firstrow = i + 1
            
            'If they are the same then add to the total stock volume for that ticker name
            Else
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
            
            End If
            
        Next i
            
        'Colour formatting for yearlychange and percentchange
        'Loop through all rows of the summary table
        For i = 2 To lastrow
                    
            'if the yearlychange value is less than 0 then highlight in red. Otherwise...
            If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
                        
            '... highlight in green
            Else
                        
            ws.Cells(i, 10).Interior.ColorIndex = 4
                        
            End If
                        
        Next i
        
           'Conditional formatting for percent change
            For i = 2 To lastrow
                If ws.Cells(i, 11).Value < 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 3
                            
                '... highlight in green
                Else
                            
                ws.Cells(i, 11).Interior.ColorIndex = 4
                            
                End If
                
            Next i
        
        'Collect the maximum decrease, maximum increase, and greatest stock volume
        'Loop through all rows of the summary table
        For i = 2 To lastrow
            
            'check if the percent change is higher than the previous row. If not...
            If ws.Cells(i, 11).Value > percentincrease Then
                
                'Create new ticker name and set variable
                ticker1 = ws.Cells(i, 9).Value
                    
                'Update the largest percent increase
                percentincrease = ws.Cells(i, 11).Value
                    
            End If
            
            'check if the percent change is lower than the previous row. If not...
            If ws.Cells(i, 11).Value < percentdecrease Then
                
                'Create new ticker name and set variable
                ticker2 = ws.Cells(i, 9).Value
                        
                'Update the largest percent decrease
                percentdecrease = ws.Cells(i, 11).Value
                        
            End If
                
            'check if the total stock volume is higher than the previous row. If not...
            If ws.Cells(i, 12).Value > stockvol Then
                
                'Create new ticker name and set variable
                stockvol = ws.Cells(i, 12).Value
                    
                'Update the largest total stock volume
                ticker3 = ws.Cells(i, 9).Value
                
            End If
            
        Next i
        
        'Update the new variables in the table
        ws.Cells(2, 16).Value = FormatPercent(percentincrease)
        ws.Cells(2, 15).Value = ticker1
        ws.Cells(3, 16).Value = FormatPercent(percentdecrease)
        ws.Cells(3, 15).Value = ticker2
        ws.Cells(4, 16).Value = stockvol
        ws.Cells(4, 15).Value = ticker3
        
    Next ws

End Sub
                                                                                                                                                                                                         