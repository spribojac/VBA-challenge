Attribute VB_Name = "Module1"
Sub stocktracker()

    'declare data variables
    Dim tickername As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalstockvolume As LongLong
    Dim openvalue As Double
    Dim closevalue As Double
    
    'declare worksheet variables
    Dim ws As Worksheet
    Dim lastrow As Long
    
    'Dim summarytabblerow As Integer
    Dim summarytablerow As Integer
    
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
        
        'allow each ticker a new location in the summary table
        summarytablerow = 2
        
        'set values to hold yearly change, percent change, total stock volume, open value and close value
        yearlychange = 0
        percentchange = 0
        totalstockvolume = 0
        firstrow = 2
        
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
        
    Next ws
        
End Sub
osoft.com *.skype.com *.teams.microsoft.us local.teams.office.com teams.microsoftonline.cn *.powerapps.com *.yammer.com *.officeapps.live.com *.office.com *.stream.azure-test.net *.microsoftstream.com *.dynamics.com *.microsoft.com onedrive.live.com *.onedrive.live.com securebroker.sharepointonline.com;
SPRequestDuration: 9
SPIisLatency: 0
MicrosoftSharePointTeamServices: 16.0.0.23926
X-Content-Type-Options: nosniff
X-MS-InvokeApp: 1; RequireReadOnly
X-AspNet-Version: 4.0.30319
X-Powered-By: ASP.NET