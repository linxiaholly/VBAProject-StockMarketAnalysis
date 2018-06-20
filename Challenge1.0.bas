Sub Stock_Analysis()

For Each ws In Worksheets

'Sunmmary Table
Dim Summary_Row As Double
Summary_Row = 2
ws.Cells(Summary_Row, 10) = ws.Cells(2, 1)
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 10) = "<ticker>"
ws.Cells(1, 11) = "<Yearly_Change>"
ws.Cells(1, 12) = "<Percent_Change>"
ws.Cells(1, 13) = "<Total_Vol>"

'Total volume
For i = 2 To LastRow

If ws.Cells(Summary_Row, 10) = ws.Cells(i, 1) Then
ws.Cells(Summary_Row, 13) = ws.Cells(Summary_Row, 13) + ws.Cells(i, 7)

ElseIf IsEmpty(ws.Cells(i, 1)) = False Then
Summary_Row = Summary_Row + 1
ws.Cells(Summary_Row, 10) = ws.Cells(i, 1)
ws.Cells(Summary_Row, 13) = ws.Cells(Summary_Row, 13) + ws.Cells(i, 7)
Else
Exit For
End If

Next i

'Min Max Date of each ticker/ open and close price / yearly change

Dim l As Double
l = 2
Dim m As Double
m = 2
Dim DateMin As Double
Dim DateMax As Double
Dim OpenPrice As Double
Dim ClosePrice As Double


For k = 2 To LastRow

If ((ws.Cells(l, 10) = ws.Cells(k, 1) And ws.Cells(k, 2) < ws.Cells(k + 1, 2))) Then
DateMin = ws.Cells(k, 2)
OpenPrice = ws.Cells(k, 3)
l = l + 1

ElseIf ((ws.Cells(m, 10) = ws.Cells(k, 1) And ws.Cells(k, 2) > ws.Cells(k + 1, 2))) Then
DateMax = ws.Cells(k, 2)
ClosePrice = ws.Cells(k, 6)

PriceChange = ClosePrice - OpenPrice
ws.Cells(m, 11) = PriceChange

On Error Resume Next
PercentChange = (PriceChange / OpenPrice)
ws.Cells(m, 12) = PercentChange

m = m + 1

End If

Next k

'Format yearly change
LastRow_Summary = ws.Cells(Rows.Count, 10).End(xlUp).Row

For n = 2 To LastRow_Summary

If ws.Cells(n, 11) >= 0 Then
ws.Cells(n, 11).Interior.ColorIndex = 4

Else
ws.Cells(n, 11).Interior.ColorIndex = 3

End If

ws.Cells(n, 12).Style = "Percent"

Next n

'Hard Question Greatest value
ws.Cells(1, 17) = "Ticker"
ws.Cells(1, 18) = "Value"
ws.Cells(2, 16) = "Greatest % Increase"
ws.Cells(3, 16) = "Greatest % Decrease"
ws.Cells(4, 16) = "Greatest Total Volume"

Dim Greatest_Inc As Double
Dim Greateast_Dec As Double
Dim Greatest_Vol As Double


Greatest_Inc = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow_Summary))
Greatest_Dec = Application.WorksheetFunction.Min(ws.Range("L2:L" & LastRow_Summary))
Greatest_Vol = Application.WorksheetFunction.Max(ws.Range("M2:M" & LastRow_Summary))

ws.Cells(2, 18) = Greatest_Inc
ws.Cells(2, 18).Style = "Percent"
ws.Cells(3, 18) = Greatest_Dec
ws.Cells(3, 18).Style = "Percent"
ws.Cells(4, 18) = Greatest_Vol

For a = 2 To LastRow_Summary

If ws.Cells(a, 12) = Greatest_Inc Then
ws.Cells(2, 17) = ws.Cells(a, 10)

ElseIf ws.Cells(a, 12) = Greatest_Dec Then
ws.Cells(3, 17) = ws.Cells(a, 10)

ElseIf ws.Cells(a, 13) = Greatest_Vol Then
ws.Cells(4, 17) = ws.Cells(a, 10)

End If
Next a

Next ws

End Sub
