Attribute VB_Name = "Module1"
'1. track change of tickers
'2. track total stock volume (see credit card excercise)
'3. track yearly change (change of opening price @ beginning of year to closing price @ end of year)
  '(color cell according to positive or negative)
'4. track percent change (relative to yearly change) ( % change = change in percent / initial percent  *  100 )
'Hint : use difference operator [<>] last row formula from lecture

 
 Sub stockData():
 
    ' create variable to track what worksheet we areon
    Dim ws As Worksheet
    
    
    
    'loops through each worksheet
    For Each ws In Worksheets
    
    'variable to track amount of rows in a worksheet
     lastrow = Cells(ws.Rows.Count, 1).End(xlUp).Row
     
       'create variable to track stock volume
        Dim totalVol As Long
        totalVol = 0
     
        'create loop that loops through repective rows/ranges to check
        For i = 2 To lastrow

        
        'create variables to track ticker value of current + next cell
        currentTick = Range("B", i)
        nextTick = Range("B", i + 1)
        
        'populates new ticker cell
        Range("I", i) = currentTick
        
        'add volume of current row to total volume
        totalVol = totalVol + Range("G", i).Value
        Range("L", i) = totalVol
        
        
        
        'check for change in ticker
        If currentTick <> nextTick Then
        totalVol = 0
        
        
        
        
        
        'end of second loop
        Next i
     
    'end of first loop
    Next ws
    
    

End Sub
