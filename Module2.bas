Attribute VB_Name = "Module2"
Sub conditionalformatting()

Dim lastrow As Long
Dim summarytablerow As Integer
Dim ws As Worksheet

    'loop through all worksheets
    For Each ws In Worksheets
    
   
        'determine the lastrow
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
            
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
        
    Next ws
    
End Sub
 ���(\O>@      � �@  `         ��G    L�G         �>@       �>@  ���(\O>@  R���Q>@      �7A  `         ��G    j�G    �p=
�c>@  =
ףp�>@  �p=
�c>@  =
ףp�>@       �@