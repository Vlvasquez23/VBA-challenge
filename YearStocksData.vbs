Sub YearlyStockData():

'Loop through all stocks for one year
'Ticker Symbol
'Yearly change from opening price to closing price
'Total Stock Volume of the stock
'conditional formatting to highligh positive change in green and negative in red.

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'Current row
        Dim i As Long
        'Start row of ticker block
        Dim j As Long
        'Index counter to fill Ticker row
        Dim Ticker_Count As Long
        'Last row column A
        Dim LastRow_A As Long
        'last row column I
        Dim LastRow_I As Long
        'Variable for percent change calculation
        Dim Percent_Change As Double
        'Variable for greatest increase calculation
        Dim Great_Incr As Double
        'Variable for greatest decrease calculation
        Dim Great_Decr As Double
        'Variable for greatest total volume
        Dim Great_Tot_Vol As Double
        
        'Get the WorksheetName
        WorksheetName = ws.Name
        
        'Create column headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'Set Ticker Counter to first row
        Ticker_Count = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in column A
        LastRow_A = Cells(Rows.Count, 1).End(xlUp).Row
               
            'Loop through all rows
            For i = 2 To LastRow_A
            
            
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Cells(Ticker_Count, 9).Value = Cells(i, 1).Value
                Cells(Ticker_Count, 10).Value = Cells(i, 6).Value - Cells(j, 3).Value
                
                If Cells(Ticker_Count, 10).Value < 0 Then
                
                Cells(Ticker_Count, 10).Interior.Color = vbRed
                
                Else

                Cells(Ticker_Count, 10).Interior.Color = vbGreen
                
                End If

                If Cells(j, 3).Value <> 0 Then
                Percent_Change = ((Cells(i, 6).Value - Cells(j, 3).Value) / Cells(j, 3).Value)
                Cells(Ticker_Count, 11).Value = Format(Percent_Change, "Percent")
                    
                Else
                    
                Cells(Ticker_Count, 11).Value = Format(0, "Percent")
                    
                End If

                Cells(Ticker_Count, 12).Value = WorksheetFunction.Sum(Range(Cells(j, 7), Cells(i, 7)))

                Ticker_Count = Ticker_Count + 1

                j = i + 1
                
                End If
            
            Next i
            
'Bonus Calculate "Greatest % Increase", "Greatest % Decrease and "Greatest Total Volume
        
        'Find last non-blank cell in column I
        LastRow_I = Cells(Rows.Count, 9).End(xlUp).Row
              
        Great_Tot_Vol = Cells(2, 12).Value
        Great_Incr = Cells(2, 11).Value
        Great_Decr = Cells(2, 11).Value
        
        
            For i = 2 To LastRow_I
            
                'For greatest total volume
                If Cells(i, 12).Value > Great_Tot_Vol Then
                Great_Tot_Vol = Cells(i, 12).Value
                Cells(4, 16).Value = Cells(i, 9).Value
                
                Else
                
                Great_Tot_Vol = Great_Tot_Vol
                
                End If
                
                'For greatest increase
                If Cells(i, 11).Value > Great_Incr Then
                Great_Incr = Cells(i, 11).Value
                Cells(2, 16).Value = Cells(i, 9).Value
                
                Else
                
                Great_Incr = Great_Incr
                
                End If
                
                'For greatest decrease
                If Cells(i, 11).Value < Great_Decr Then
                Great_Decr = Cells(i, 11).Value
                Cells(3, 16).Value = Cells(i, 9).Value
                
                Else
                
                Great_Decr = Great_Decr
                
                End If
                
            'Summarize results in Cells
            Range("Q2").Value = Format(Great_Incr, "Percent")
            Range("Q3").Value = Format(Great_Decr, "Percent")
            Range("Q4").Value = Format(Great_Tot_Vol, "Scientific")
            
            Next i
            
        'Auto column width
        Worksheets(WorksheetName).Columns("I:Q").AutoFit
            
    Next ws
        
End Sub
