Sub Ticker()

Dim ws As Worksheet
Dim openprice As Double
Dim closeprice As Double
Dim totalstockvol As Double

'loop through the sheets in the workbook
For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "percent change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    startingrow = 2
    openprice = ws.Range("C2").Value

'find the last row
For lastrow = 2 To ws.Range("A2").End(xlDown).Row
    
    'find the edge
    If ws.Cells(lastrow, 1).Value <> ws.Cells(lastrow + 1, 1).Value Then
            
        'Ticker
        ws.Cells(startingrow, 9).Value = ws.Cells(lastrow, 1).Value
        
        'calculate the total stock vol
        totalstockvol = totalstockvol + ws.Cells(lastrow, 7).Value
        
        'get closing price
        closeprice = ws.Cells(lastrow, 6).Value
            
        'yearly change in price
        ws.Cells(startingrow, 10).Value = closeprice - openprice
        
        'add conditional formating
        If ws.Cells(startingrow, 10) >= 0 Then
            ws.Cells(startingrow, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(startingrow, 10).Interior.ColorIndex = 3
        End If
        
        'percent change
        If openprice = 0 Then
            ws.Cells(startingrow, 11).Value = "Not a number"
        Else
            ws.Cells(startingrow, 11).Value = (closeprice - openprice) / openprice
            ws.Cells(startingrow, 11).NumberFormat = "0.00%"
        End If
        
        'total stock volume
        ws.Cells(startingrow, 12).Value = totalstockvol
                      
        'go to the the next group(ticker) and reset
        startingrow = startingrow + 1
            
        'rest total stock volume
        totalstockvol = 0
        
        'new openprice
        openprice = ws.Cells(lastrow + 1, 3).Value
            
    Else
        'calculate the total stock vol
        totalstockvol = totalstockvol + ws.Cells(lastrow, 7).Value
    End If
        
Next lastrow

'determining greatest increase, decrease, and total volume

    Greatestincrease = ws.Application.WorksheetFunction.Max(ws.Range("K:K"))
    GreatestTotVol = ws.Application.WorksheetFunction.Max(ws.Range("L:L"))
    Greatestdecrease = ws.Application.WorksheetFunction.Min(ws.Range("K:K"))
        
    For i = 2 To ws.Range("i2").End(xlDown).Row
            
        If ws.Cells(i, 11).Value = Greatestincrease Then
            ws.Range("p2") = Greatestincrease
            ws.Range("p2").NumberFormat = "0.00%"
            ws.Range("o2") = ws.Cells(i, 9).Value
    
        
        ElseIf ws.Cells(i, 11).Value = Greatestdecrease Then
        ws.Range("p3") = Greatestdecrease
        ws.Range("p3").NumberFormat = "0.00%"
        ws.Range("o3") = ws.Cells(i, 9).Value
    
        
        ElseIf ws.Cells(i, 12).Value = GreatestTotVol Then
        ws.Range("p4") = GreatestTotVol
        ws.Range("o4") = ws.Cells(i, 9).Value
    
        End If
    
    Next i

Next ws

End Sub