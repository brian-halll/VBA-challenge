Attribute VB_Name = "Module1"
'1. track change of tickers
'2. track total stock volume (see credit card excercise)
'3. track yearly change (change of opening price @ beginning of year to closing price @ end of year)
  '(color cell according to positive or negative)
'4. track percent change (relative to yearly change) ( % change = change in percent / initial percent  *  100 )
'Hint : use difference operator [<>] last row formula from lecture

 Sub stockData():
 
    ' create variable to track what worksheet we areon
    ' Dim ws As Worksheet
    
    
    'loops through each worksheet
    For Each ws In Worksheets
    
    'variable to track amount of rows in a worksheet
     lastrow = Cells(ws.Rows.Count, 1).End(xlUp).Row
     
     ws.Range("i1").Value = "Ticker"
     ws.Range("l1").Value = "Total Stock Volume"
     ws.Range("j1").Value = "Yearly  Change"
     ws.Range("k1").Value = "Percent Change"
     
     
     
       'create variable to track stock volume
        Dim totalVol As LongLong
        
        
        
        'record opening price of initial stock in a worksheet
        openPrice = ws.Cells(2, 3).Value
        
     
        'create loop that loops through repective rows/ranges to check
        For i = 2 To lastrow

        
        'create variables to track ticker value of current + next cell
        currentTick = ws.Cells(i, 1).Value
        nextTick = ws.Cells(i + 1, 1).Value
        
        'populates new ticker cell
        ws.Cells(i, 9).Value = currentTick
        
        'add volume of current row to total volume
        totalVol = totalVol + ws.Cells(i, 7).Value
        ws.Cells(i, 12) = totalVol
        
        
        
        'check for change in ticker and update values accordingly
        If currentTick <> nextTick Then
        totalVol = 0
        openPrice = ws.Cells(i + 1, 3).Value
        
        Else
        
        
        End If
        
        
        'end of second loop
        Next i
     
    'end of first loop
    Next ws
    
    

End Sub
