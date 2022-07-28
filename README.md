# vba-Challange-
The file uploded is a empty worksheet with the marcos names "stocks" was used as the script for Multiple_year_stock_data
If the script does not open for whatever reason here is a copy of the script 

Sub Stocks
Dim outputrow As Integer
Dim yr_open As Double
Dim yr_close As Double
Dim yr_change as Double
Dim totalstockvol As Double


outputrow= 2


For I = 2 to 753001

If Cells(i, 1).value <> Cells(I+ 1).value Then
	Cells(outputrow, 9).value = Cells(I,1). value
	yr_open = Cells(2, 3).value
	yr_close= Cells(I, 6).value
	yr_change = yr_close- yr_open
	Cells(outputrow, 10).value= yr_change
	Cells(outputrow, 11).value= (yr_change / yr_open)
	totalstockvol = totalstockvol + Cells(I, 7).value
	cells(outputrow, 12).value = totalstockvol
	yr_open= Cells(2 + I, 3).value
	outputrow= outputrow + 1 

End If 
If IsEmpty(Cells(I, 10). value) = True Then
	Cells(I, 10).Interior.ColorIndex = 2
ElseIf Cells(I, 10). value > 0 Then
	Cells(I, 10).Interior.ColorIndex = 4
ElseIf Cells(I, 10).value <= 0 Then
	Cells(I, 10). Interior. ColorIndex = 3


End If


Next I
End Sub 
	





