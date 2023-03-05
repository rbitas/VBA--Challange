Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
      
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    Dim Pricepercentchange As Double
    Dim ticker As String
    ticker = " "
    Dim Ticker_symbol As Double
    Ticker_symbol = 0
    

    Dim lastrow As Long
    Dim i As Long
    Dim j As Long
        j = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        
    
     Dim TickerRow As Long
        TickerRow = 1
        
     For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TickerRow = TickerRow + 1
        ticker = ws.Cells(i, 1).Value
        ws.Cells(TickerRow, "I").Value = ticker
    
        ws.Cells(TickerRow, "J").Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
            If ws.Cells(TickerRow, "J").Value < 0 Then
            ws.Cells(TickerRow, "J").Interior.ColorIndex = 3
            
            Else
            ws.Cells(TickerRow, "J").Interior.ColorIndex = 4
            
            End If
            
        If ws.Cells(j, 3).Value <> 0 Then
        
            Pricechangepercent = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
            
            ws.Cells(TickerRow, 11).Value = Format(Pricechangepercent, "Percent")
            
            Else
            
            ws.Cells(TickerRow, 11).Value = Format(0, "Percent")
            
            End If
            
        ws.Cells(TickerRow, 12).Value = Application.WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        
        j = i + 1

            
        End If
        
    Next i
   
lastrowforI = ws.Cells(Rows.Count, 9).End(xlUp).Row
 
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

percentchange = ws.Cells(i, 11).Value
volume = ws.Cells(i, 12).Value

For i = 2 To lastrowforI

    If percentchange > GreatestIncrease Then
    GreatestIncrease = percentchange
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = GreatestIncrease
    
    End If
    
    If percentchange < GreatestDecrease Then
    GreatestDecrease = percentchange
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = GreatestDecrease
    End If
    
    If volume > GreatestVolume Then
    GreatestVolume = volume
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = GreatestVolume
               
     End If
     
     
Next i


Next ws
            
End Sub

