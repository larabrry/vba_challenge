Sub allSheets()
    Dim ws As Worksheet
    
    
    For Each ws In ThisWorkbook.Worksheets
    
        Call stock(ws)
        
    Next ws

End Sub


Sub stock(ws)
'Create the column headings

    [I1:L1] = [{"Ticker", "Yearly Change", "Percent Change", "Total Stock Volume"}]
    [P1:Q1] = [{"Ticker", "Value"}]
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    

'Define variable
Dim lastrow As Long          'The last row in the stock data.
Dim nextSumRow As Long       'The next open row on the summary table.
Dim i As Long                'The iteration variable in For loop.
Dim strTicker As String      'Ticker name.
Dim dblYearOpen As Double    'The price the stock opened the year at.
Dim dblYearClose As Double   'The price the stock closed the year at.
Dim dblStockVol As Double    'The volume of stock.

Dim dblGreatestPercInc As Double
Dim dblGreatestPerDec As Double
Dim dblGreatestTotalVol As Double

Dim dblGreatestPercIncTick As String
Dim dblGreatestPerDecTick As String
Dim dblGreatestTotalVolTick  As String

Dim tmpPercChange As Double

dblGreatestPercInc = 0
dblGreatestPerDec = 0
dblGreatestTotalVol = 0

dblGreatestPercIncTick = ""
dblGreatestPerDecTick = ""
dblGreatestTotalVolTick = ""


dblStockVol = 0
'determine last row
lastrow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

strTicker = Range("A2").Value
dblYearOpen = Range("C2").Value

nextSumRow = 2

tmpPercChange = 0
'start loop
For i = 2 To lastrow + 1
    
    If ws.Range("A" & i).Value <> strTicker Then
    
        'add the previous ticker's close price and calculate percent change
        dblYearClose = Range("F" & i - 1).Value
        
        tmpPercChange = Round(-(1 - (dblYearClose / dblYearOpen)) * 100, 2)
  
        'find greatest % inc, greastest % dec, greatest total vol
        If tmpPercChange > 0 Then
            If dblGreatestPercInc < tmpPercChange Then
                dblGreatestPercInc = tmpPercChange
                dblGreatestPercIncTick = strTicker
            End If
        End If
        If tmpPercChange < 0 Then
            If dblGreatestPerDec > tmpPercChange Then
                dblGreatestPerDec = tmpPercChange
                dblGreatestPerDecTick = strTicker
                
            End If
        End If
        
        If dblGreatestTotalVol < dblStockVol Then
            dblGreatestTotalVol = dblStockVol
            dblGreatestTotalVolTick = strTicker
        End If

        'print the previous ticker name and its yearly change value.
        ws.Range("I" & nextSumRow).Value = strTicker
        ws.Range("J" & nextSumRow).Value = Round(dblYearClose - dblYearOpen, 2)
            'apply conditional formatting
            If Round(dblYearClose - dblYearOpen, 2) < 0 Then
                ws.Range("J" & nextSumRow).Interior.Color = RGB(255, 0, 0)
            Else
                ws.Range("J" & nextSumRow).Interior.Color = RGB(0, 255, 0)
            End If
        ws.Range("K" & nextSumRow).Value = tmpPercChange
        ws.Range("L" & nextSumRow).Value = dblStockVol
        'Iterate the summary row and setup new ticker data.
        nextSumRow = nextSumRow + 1
        strTicker = ws.Range("A" & i).Value
        dblYearOpen = ws.Range("C" & i).Value
        dblStockVol = 0
        
    End If
        'if condition is false, just add up the previous tickers volume to column L
        dblStockVol = dblStockVol + ws.Range("G" & i).Value
Next i

'print greatest % inc,dec, and total vol in column P and Q
ws.Range("P2").Value = dblGreatestPercIncTick
ws.Range("Q2").Value = dblGreatestPercInc
 
ws.Range("P3").Value = dblGreatestPerDecTick
ws.Range("Q3").Value = dblGreatestPerDec
 
ws.Range("P4").Value = dblGreatestTotalVolTick
ws.Range("Q4").Value = dblGreatestTotalVol

'Add % sign to yearly change values
ws.Range("Q2:Q3").NumberFormat = "0.00%"


End Sub
