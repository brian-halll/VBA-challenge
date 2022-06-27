Attribute VB_Name = "Module1"
'1. track change of tickers
'2. track total stock volume (see credit card excercise)
'3. track yearly change (change of opening price @ beginning of year to closing price @ end of year)
  '(color cell according to positive or negative)
'4. track percent change (relative to yearly change) ( % change = change in percent / initial percent  *  100 )
'Hint : use difference operator [<>] last row formula from lecture

 Sub stockData():
 

    Dim percentChange As Double
    Dim yearlyChange As Double
    Dim totalVol As LongLong
    Dim greatestVol As LongLong
    greatestVol = 0
    
    Dim greastestIncrease As Double
    greastestIncrease = 0
    
    Dim greastestDecrease As Double
    greatestDecrease = 0

    'loops through each worksheet
    For Each ws In Worksheets
    
    'variable to track amount of rows in a worksheet
     lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
     
     ws.Range("i1").Value = "Ticker"
     ws.Range("l1").Value = "Total Stock Volume"
     ws.Range("j1").Value = "Yearly  Change"
     ws.Range("k1").Value = "Percent Change"
     

        'record opening price of initial stock in a worksheet
        openPrice = ws.Cells(2, 3).Value
        
     
        'create loop that loops through repective rows/ranges to check
        For i = 2 To lastrow

        
        'create variables to track ticker value of current + next cell
        currentTick = ws.Cells(i, 1).Value
        nextTick = ws.Cells(i + 1, 1).Value
        
        'populates new ticker cell
        ws.Cells(i, 9).Value = currentTick
        
        'add volume of current row to total volume and populate cell
        totalVol = totalVol + ws.Cells(i, 7).Value
        ws.Cells(i, 12) = totalVol
        
        'calaculate yearly change and populate cell
        yearlyChange = ws.Cells(i, 6).Value - openPrice
        ws.Cells(i, 10).Value = yearlyChange
        

        
        'color cell according positive or negative yearly change
        If yearlyChange < 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
            Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
        End If
        
        
        'calculate percent change (change in percent / initial percent  *  100 ) and populate cell , also check for  greastest dcrease and increases
        percentChange = ((ws.Cells(i, 6).Value - openPrice) / openPrice)
        If percentChange > greatestIncrease Then
            increaseTick = currentTick
            greatestIncrease = percentChange
            
            ElseIf percentChange < greastestDecrease Then
            decreaseTick = currentTick
            greastestDecrease = percentChange
            
            End If
        
        ws.Cells(i, 11).Value = Format(percentChange, "##0.00%")
        
        'check for greastest total volume
         If totalVol > greatestVol Then
                greatestVol = totalVol
                volTick = currentTick
                
                Else
                
                End If
        
        
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
    
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest %  Increase"
    Cells(3, 14).Value = "Greatest %  Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(4, 15).Value = volTick
    Cells(3, 15).Value = increaseTick
    Cells(2, 15).Value = decreaseTick
    Cells(4, 16).Value = greatestVol
    Cells(3, 16).Value = Format(greatestIncrease, "##0.00%")
    Cells(2, 16).Value = Format(greastestDecrease, "##0.00%")
    
    
    

End Sub
