Sub ticker():

For Each WS In Worksheets

WS.Cells(1, 9).Value = "Ticker"
WS.Cells(1, 10).Value = "Yearly Change"
WS.Cells(1, 11).Value = "Percent Change"
WS.Cells(1, 12).Value = "Total Stock Volume"

Dim tick, tick_max1, tick_max2, tick_max3 As String
Dim Rows As Integer
Dim op_price, cl_price, year_chng, perc_chng As Double
Dim stock, Grt_incrs, Grt_dec, Total_vol, row_num1, row_num2, row_num3 As Double

stock = 0
Rows = 2
Grt_incrs = 0
Grt_dec = 0
Total_vol = 0
row_limit = WS.Cells(2, 1).End(xlDown).Row
op_price = WS.Cells(2, 3).Value

For I = 2 To row_limit

If WS.Cells(I, 1).Value <> WS.Cells(I + 1, 1).Value Then

tick = WS.Cells(I, 1).Value
WS.Range("I" & Rows).Value = tick

cl_price = WS.Cells(I, 6).Value

year_chng = cl_price - op_price

perc_chng = (year_chng / op_price) * 100

op_price = WS.Cells(I + 1, 3).Value

WS.Range("J" & Rows).Value = year_chng

perc_chng = Round(perc_chng, 2)

WS.Range("K" & Rows).Value = "%" & perc_chng

If perc_chng > Grt_incrs Then
Grt_incrs = perc_chng
row_num1 = Rows
tick_max1 = WS.Cells(row_num1, 9).Value


End If

If perc_chng < Grt_dec Then
Grt_dec = perc_chng
row_num2 = Rows
tick_max2 = WS.Cells(row_num2, 9).Value

End If



If year_chng < 0 Then
WS.Range("J" & Rows).Interior.ColorIndex = 3

Else
WS.Range("J" & Rows).Interior.ColorIndex = 4
End If

stock = stock + WS.Cells(I, 7).Value

WS.Range("L" & Rows).Value = stock

Rows = Rows + 1

stock = 0

Else

stock = stock + WS.Cells(I, 7).Value


If stock > Total_vol Then
Total_vol = stock
row_num3 = Rows
tick_max3 = WS.Cells(row_num3, 9).Value


End If

End If


Next I



WS.Cells(2, 15).Value = "Greatest % Increase"
WS.Cells(3, 15).Value = "Greatest % Decrease"
WS.Cells(4, 15).Value = "Greatest Total Volume"
WS.Cells(1, 16).Value = "Ticker"
WS.Cells(1, 17).Value = "Value"
WS.Cells(2, 16).Value = tick_max1
WS.Cells(2, 17).Value = "%" & Grt_incrs

WS.Cells(3, 16).Value = tick_max2
WS.Cells(3, 17).Value = "%" & Grt_dec

WS.Cells(4, 16).Value = tick_max3
WS.Cells(4, 17).Value = Total_vol

Next WS


End Sub
 

