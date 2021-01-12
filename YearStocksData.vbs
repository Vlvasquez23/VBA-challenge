Sub YearStockData():

' Loop for All worksheets(ws)

For Each ws In Worksheets

    Dim Worksheetname As String
    Dim i As Long
    Dim j As Long
    Dim Total As Double
    
    
    
    
    
    
' Column Headers/Data Labels
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Greatest % Increase"
    ws.Range("O2").Value = "Greatest % Decrease"
    ws.Range("O3").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
'Determine Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
    
   
        
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         
         
         
    
    
  Next
    
    
    
    
End Sub

