Attribute VB_Name = "Module1"
Sub stockchange()


For Each ws In Worksheets


'init variables
Dim ticker As String
Dim total_vol As LongLong
Dim iter As Long
Dim lastrow As Long
Dim year_start As Double
Dim year_end As Double
Dim year_change As Double
Dim percent As Double
Dim percent_increase As Double
Dim percent_decrease As Double
Dim max_vol As LongLong



total_vol = 0
iter = 2
year_start = 0
year_end = 0
year_change = 0
percent = 0
percent_increase = 0
percent_decrease = 0
max_vol = 0


'count to last row of sheet
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'set headers
ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total stock volume"
ws.Range("o1").Value = "Ticker"
ws.Range("p1").Value = "Value"
ws.Range("n2").Value = "Greatest % increase"
ws.Range("n3").Value = "Greatest % decrease"
ws.Range("n4").Value = "Greatest total volume"



'start loop
For i = 2 To lastrow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'add current cell to vol, set ticker name, and year end
total_vol = total_vol + ws.Cells(i, 7)
ticker = ws.Cells(i, 1)
year_end = ws.Cells(i, 6)

'find year change and % change
year_change = (year_end - year_start)
percent = (year_change / year_start)

'compare to max and min change
If percent > percent_increase Then
percent_increase = percent
ws.Range("p2").Value = percent_increase
ws.Range("p2").NumberFormat = "0.00%"
ws.Range("o2").Value = ticker


ElseIf percent < percent_decrease Then
percent_decrease = percent
ws.Range("p3").Value = percent_decrease
ws.Range("p3").NumberFormat = "0.00%"
ws.Range("o3").Value = ticker


End If

'find max volume
If total_vol > max_vol Then
max_vol = total_vol
ws.Range("p4").Value = max_vol
ws.Range("o4").Value = ticker
End If

'set values in summary table
ws.Range("i" & iter).Value = ticker
ws.Range("l" & iter).Value = total_vol
ws.Range("j" & iter).Value = Round(year_change, 2)
ws.Range("k" & iter).Value = Round(percent, 4)

'format percent
ws.Range("k" & iter).NumberFormat = "0.00%"

'format negative red
If ws.Range("j" & iter).Value <= 0 Then
ws.Range("j" & iter).Interior.ColorIndex = 3

'format positive green
ElseIf ws.Range("j" & iter).Value > 0 Then
ws.Range("j" & iter).Interior.ColorIndex = 4

End If




'iter one through summary table
iter = iter + 1

'reset vars
total_vol = 0
year_start = 0
year_end = 0
percent = 0
year_change = 0


'if cell above is different set year start
ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
year_start = ws.Cells(i, 3).Value



Else
total_vol = total_vol + ws.Cells(i, 7).Value

End If




Next i

Next ws



End Sub
