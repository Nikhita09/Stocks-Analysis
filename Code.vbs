Sub AllStockAnalysis()

yearvalue = InputBox("What year would you like to run the analysis on?")

Worksheets("All Stock Analysis").Activate

Range("A1") = "All Stocks (" + yearvalue + ")"

'Create a header row
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Starting Price"
Cells(3, 4).Value = "Ending Price"
Cells(3, 5).Value = "Return"

'Set a tickerIndex as 0. Increase if next row's ticker doesnt match
Dim tickers() As String
    tickerindex = 0
ReDim tickers(tickerindex)
    
'Create arrays for all volumes, starting and ending price
Dim totalVolume() As Long
Dim startingPrice() As Double
Dim endingPrice() As Double
ReDim tickers(tickerindex)
ReDim totalVolume(tickerindex)
ReDim startingPrice(tickerindex)
ReDim endingPrice(tickerindex)


Worksheets(yearvalue).Activate

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

i = 0
    


        For j = 2 To RowCount
            
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
                Worksheets(yearvalue).Activate
                totalVolume(tickerindex) = totalVolume(tickerindex) + Cells(j, 8).Value
                endingPrice(tickerindex) = Cells(j, 6).Value
                
    Worksheets("All Stock Analysis").Activate
        Cells(4 + i, 1).Value = tickers(tickerindex)
        Cells(4 + i, 2).Value = totalVolume(tickerindex)
        Cells(4 + i, 3).Value = startingPrice(tickerindex)
        Cells(4 + i, 4).Value = endingPrice(tickerindex)
        Cells(4 + i, 5).Value = (endingPrice(tickerindex) / startingPrice(tickerindex)) - 1
                
i = i + 1
Worksheets(yearvalue).Activate
tickerindex = tickerindex + 1

ReDim tickers(tickerindex)
ReDim totalVolume(tickerindex)
ReDim startingPrice(tickerindex)
ReDim endingPrice(tickerindex)


            ElseIf Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
                tickers(tickerindex) = Cells(j, 1).Value
                totalVolume(tickerindex) = totalVolume(tickerindex) + Cells(j, 8).Value
                startingPrice(tickerindex) = Cells(j, 6).Value
            
            Else
                totalVolume(tickerindex) = totalVolume(tickerindex) + Cells(j, 8).Value
            End If
             
        Next j
        
Worksheets("All Stock Analysis").Activate
    Range("A3:E3").Font.FontStyle = "Bold"
    Range("A3:E3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    lrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Range("B4:B" & lrow).NumberFormat = "#,##0"
    Range("C4:C" & lrow).NumberFormat = "0.00"
    Range("D4:D" & lrow).NumberFormat = "0.00"
    Range("E4:E" & lrow).NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    For i = 4 To lrow
    If Cells(i, 5) < 0 Then
            'Color the cell red
            Cells(i, 5).Interior.Color = vbRed
      
        ElseIf Cells(i, 5) > 0 Then
            'Color the cell green
            Cells(i, 5).Interior.Color = vbGreen
      
        Else
            'Clear the cell color
            Cells(i, 5).Interior.Color = xlNone
      
        End If
  
    Next i

End Sub
