Sub columns():
Dim n As Integer
For n = 1 To 3
    'create column titles'
    Worksheets(n).Range("I1").Value = "Ticker"
    Worksheets(n).Range("J1").Value = "Yearly Change"
    Worksheets(n).Range("K1").Value = "Percent Change"
    Worksheets(n).Range("L1").Value = "Total Stock Volume"
    Worksheets(n).Range("N2").Value = "Greatest Percent Increase"
    Worksheets(n).Range("N3").Value = "Greatest Percent Decrease"
    Worksheets(n).Range("N4").Value = "Greatest Total Volume"
    Worksheets(n).Range("O1").Value = "Value"
    Worksheets(n).Range("P1").Value = "Ticker"
    Next n
    
End Sub

Sub ticker_values():
For Each ws In Worksheets

Dim ticker As String
Dim ticker_row As Double
Dim yearly_row As Double
Dim percent_row As Double
Dim j As Long
Dim volume_row As Double
Dim max_percent As Double
Dim min_percent As Double
Dim max_volume As Double
Dim lastrow1 As Long
Dim lastrowYC As Long
Dim lastrowPC As Long
Dim lastrowTV As Long

'determine last rows
lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastrowYC = ws.Cells(Rows.Count, 10).End(xlUp).Row
lastrowPC = ws.Cells(Rows.Count, 11).End(xlUp).Row
lastrowTV = ws.Cells(Rows.Count, 12).End(xlUp).Row

j = 2
yearly_row = 2
percent_row = 2
ticker_row = 2
volume_row = 2
For i = 2 To lastrow1
'create row with each ticker name

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value

'yearly change
ws.Cells(yearly_row, "J").Value = (ws.Cells(i, "F") - ws.Cells(j, "C"))
    'conditional formatting so positive values are green and negative values are red
    If ws.Cells(yearly_row, "J").Value > 0 Or ws.Cells(yearly_row, "J").Value = 0 Then
    ws.Cells(yearly_row, "J").Interior.ColorIndex = 4
    Else
    ws.Cells(yearly_row, "J").Interior.ColorIndex = 3
    End If
    
'percent change
ws.Cells(percent_row, "K").Value = ((ws.Cells(i, "F") - ws.Cells(j, "C")) / (ws.Cells(j, "C")))
'format as percent
ws.Cells(percent_row, "K").NumberFormat = "0.00%"
'total volume
ws.Cells(volume_row, "L").Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

'put ticker name in new column
ws.Range("I" & ticker_row).Value = ticker

'add count to ticker row so next ticker is in a new line
ticker_row = ticker_row + 1
'add count to yearly row so next yearly change is in a new line
yearly_row = yearly_row + 1
'add count to percent row
percent_row = percent_row + 1
volume_row = volume_row + 1


j = i + 1

End If

Next i

'find greatest and least % yearly change and greatest total value
max_percent = WorksheetFunction.max(Range(ws.Cells(2, "K"), ws.Cells(lastrowPC, "K")))
ws.Cells(2, 15).Value = max_percent
ws.Cells(2, 15).NumberFormat = "0.00%"
min_percent = WorksheetFunction.Min(Range(ws.Cells(2, "K"), ws.Cells(lastrowPC, "K")))
ws.Cells(3, 15).Value = min_percent
ws.Cells(3, 15).NumberFormat = "0.00%"
max_volume = WorksheetFunction.max(Range(ws.Cells(2, 12), ws.Cells(lastrowTV, 12)))
ws.Cells(4, 15).Value = max_volume

'put ticker value for greatest yearly change
For i = 2 To lastrowYC
If ws.Cells(2, 15).Value = ws.Cells(i, "K") Then
ws.Cells(2, 16).Value = ws.Cells(i, "I").Value

End If
Next i

'put ticker value for least yearly change
For i = 2 To lastrowYC
If ws.Cells(3, 15).Value = ws.Cells(i, "K") Then
ws.Cells(3, 16).Value = ws.Cells(i, "I").Value
End If
Next i

'put ticker value for greatest total value
For i = 2 To lastrowTV
If ws.Cells(4, 15).Value = ws.Cells(i, "L") Then
ws.Cells(4, 16).Value = ws.Cells(i, "I").Value

End If
Next i

Next ws

End Sub
