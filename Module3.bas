Attribute VB_Name = "Module3"
Sub standouts()

'go through the row for column J. If the result is bigger than the one before
'store it as value and then move onto the next one
'if it is smaller then move on.
    
    'declare worksheet variables
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim percentincrease As Double
    Dim percentdecrease As Double
    Dim stockvol As LongLong
    
    'Loop over each worksheet
    For Each ws In Worksheets
    
        'Create table for values
        ws.Range("O1").EntireColumn.Insert
        ws.Cells(1, 15).Value = "Ticker"
        ws.Range("O1").EntireColumn.AutoFit
        
        ws.Range("P1").EntireColumn.Insert
        ws.Cells(1, 16).Value = "Value"
        ws.Range("P1").EntireColumn.AutoFit
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Initialise the lastrow
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
        
        'Initialise the maximum value
        percentincrease = 0
        percentdecrease = 0
        stockvol = 0
        ticker1 = 0
        ticker2 = 0
        ticker3 = 0
        
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
D a t a . D o c . L i c e n s e C a t e g o r y " :   8 ,   " D a t a . D o c . S i z e I n B y t e s " :   1 0 8 4 0 9 4 5 6 ,   " D a t a . D o c . R e a d O n l y R e a s o n s " :   0 ,   " D a t a . D o c . T e n a n t I d " :   " 9 1 8 8 0 4 0 d - 6 c 6 7 - 4 c 5 b - b 1 1 2 - 3 6 a 3 0 4 b 6 6 d a d " ,   " D a t a . D o c . I d e n t i t y T e l e m e t r y I d " :   " 0 0 0 0 0 0 0 0 - 0 0 0 0 - 0 0 0 0 - B 3 F 6 - C 7 A D A F A 9 A 0 3 5 " }   ��L    ɽ	 `��L  ��gL  p/Data%20Analyti