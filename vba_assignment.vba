Sub Stocks()
Dim outputrow As Integer
Dim yr_open As Double
Dim yr_close As Double
Dim yr_change As Double
Dim totalstockvol As Double

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

outputrow = 2



For I = 2 To 753001

If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
    Cells(outputrow, 9).Value = Cells(I, 1).Value
    yr_open = Cells(2, 3).Value
    yr_close = Cells(I, 6).Value
    yr_change = yr_close - yr_open
    Cells(outputrow, 10).Value = yr_change
    Cells(outputrow, 11).Value = (yr_change / yr_open)
    totalstockvol = totalstockvol + Cells(I, 7).Value
    Cells(outputrow, 12).Value = totalstockvol
    yr_open = Cells(2 + I, 3).Value
    outputrow = outputrow + 1

End If
If IsEmpty(Cells(I, 10).Value) = True Then
    Cells(I, 10).Interior.ColorIndex = 2
ElseIf Cells(I, 10).Value > 0 Then
    Cells(I, 10).Interior.ColorIndex = 4
ElseIf Cells(I, 10).Value <= 0 Then
    Cells(I, 10).Interior.ColorIndex = 3


End If


Next I
End Sub
