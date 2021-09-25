Attribute VB_Name = "Module1"
Sub StockAnalysis():

For Each ws In Worksheets
    
    'fill text for header cells
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'declare (and intialize some) variables
    Dim ticker As String
    
    Dim yearlyOpen As Double
    
    Dim yearlyClose As Double
    
    Dim yearlyChange As Double
    
    Dim percentChange As Double
    
    Dim opener As Double
    opener = 2
    
    Dim totalVol As Long
    totalVol = 0
    
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim summaryTableRow As Long
    summaryTableRow = 2
    
         ' loop through all of the ticker rows
        For Row = 2 To lastRow
            
            ' check to see if within the same ticker
            If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
                ' Set (reset) the ticker name
                ticker = ws.Range("A" & Row).Value
                ' Add to the volume total before the change in ticker
                 totalVol = totalVol + ws.Range("G" & Row).Value
                 
                 'holds value for yearly open at start of new ticker
                yearlyOpen = ws.Range("C" & opener)
                'holds value for yearly close at end of ticker
                yearlyClose = ws.Range("F" & Row)
                'calculate yearly change
                yearlyChange = yearlyClose - yearlyOpen
                
                    'calculate percent change
                    If yearlyOpen = 0 Then
                    percentChange = 0
                    
                    Else
                   percentChange = yearlyChange / yearlyOpen
                
                    End If
                
                'Add the values to the summary table
                'Add the ticker
                ws.Range("J" & summaryTableRow).Value = ticker
                'Add the volume total
                ws.Range("M" & summaryTableRow).Value = totalVol
                'Add the yearly change
                ws.Range("K" & summaryTableRow).Value = yearlyChange
                'Add the percent change
                ws.Range("L" & summaryTableRow).Value = percentChange
                
                
                'formatting
                If ws.Range("K" & summaryTableRow).Value >= 0 Then
                    ws.Range("K" & summaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & summaryTableRow).Interior.ColorIndex = 3
                End If
                
                ws.Columns("J:Q").AutoFit
                ws.Range("L2:L" & summaryTableRow).NumberFormat = "0.00%"
                   
                'next row on summary table
                summaryTableRow = summaryTableRow + 1
                'reset volume total
                totalVol = 0
                
            End If
        
        Next Row
        
        'find the greatest total volume
        For Row = 2 To lastRow
            
            If ws.Range("M" & Row) > ws.Range("Q4") Then
                ws.Range("P4") = ws.Range("J" & Row)
                ws.Range("Q4") = ws.Range("M" & Row)
            End If
            
        Next Row
        
    Next ws

End Sub
