Sub ticker():

For Each ws In Worksheets

'Get Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'Get Last row in each worksheet
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
'Initial Value
    numOpen = 0
    numClose = 0
    numStock = 0
    OpenCounter = 0
    x = 2 'Initial value for info
    
'Add all the same ticker values
    For i = 2 To lastRow
        OpenCounter = OpenCounter + 1 'Counts how many rows went by since the open ticker
        numStock = numStock + ws.Cells(i, 7).Value 'Counts up the volume of stocks

        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            numOpen = ws.Cells(i - OpenCounter + 1, 3).Value 'Gets the value at the start of the ticker
            numClose = ws.Cells(i, 6).Value 'Gets the value at the end of the ticker
            ws.Cells(x, 9) = ws.Cells(i, 1) 'Adds ticker to info
            ws.Cells(x, 10) = numClose - numOpen 'Yearly Change
            
            'Change color of the cell
            If ws.Cells(x, 10) < 0 Then
                    ws.Cells(x, 10).Interior.ColorIndex = 3
            Else
                    ws.Cells(x, 10).Interior.ColorIndex = 4
            End If
            
            'Situation for denominator being 0
            If numOpen <> 0 Then
                ws.Cells(x, 11) = Round(((numClose / numOpen) - 1), 4) 'Percent Change
            Else
            End If
            
            ws.Cells(x, 11).NumberFormat = "0.00%" 'Changes format to percent
            ws.Cells(x, 12) = numStock 'Total Stock Volume
            x = x + 1 'Next Line for info
            
            
            'Reset Initial Values
            numOpen = 0
            numClose = 0
            numStock = 0
            OpenCounter = 0
        End If
        
    Next i

Next ws

End Sub
