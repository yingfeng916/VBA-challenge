Sub challenge()

'Declare values
Dim b As Integer
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Find last row dynamically\
endRowTicker = Cells(Rows.Count, 9).End(xlUp).Row
rangeK = Range("K2:K" & endRowTicker)
rangeL = Range("L2:L" & endRowTicker)

Range("Q2").Value = Application.WorksheetFunction.Max(rangeK) 'Greatest % Increase
Range("Q2").NumberFormat = "0.00%"
Range("Q3").Value = Application.WorksheetFunction.Min(rangeK)  'Smallest % Decrease
Range("Q3").NumberFormat = "0.00%"
Range("Q4").Value = Application.WorksheetFunction.Max(rangeL) 'Greatest Total Volume

For b = 2 To endRowTicker
    If Range("K" & b).Value = Range("Q2").Value Then
        Range("P2").Value = Range("I" & b).Value
    ElseIf Range("K" & b).Value = Range("Q3").Value Then
        Range("P3").Value = Range("I" & b).Value
    ElseIf Range("L" & b).Value = Range("Q4").Value Then
        Range("P4").Value = Range("I" & b).Value
    End If
Next b

End Sub