Sub stock()

'Declare variables
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim ticker_row As Integer
Dim newTickerRow As Boolean

'Initialize Values
rowIncrement = 2
yearly_change = 0
total_volume = 0
newTickerRow = True

'Write headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("k1").Value = "Yearly Percentage Change"
Range("l1").Value = "Total Volume"


'Find last row dynamically
endRow = Cells(Rows.Count, 1).End(xlUp).Row

'Begin main
For a = 2 To endRow
    If Cells(a + 1, 1).Value <> Cells(a, 1).Value Then
        Range("I" & rowIncrement).Value = Cells(a, 1).Value  'Write ticker value
        Range("j" & rowIncrement).Value = Round(Cells(a, 6).Value - stored_open, 2)  'Write yearly change value
        If Range("j" & rowIncrement).Value > 0 Then 'Change colour according to +/-
            Range("j" & rowIncrement).Interior.ColorIndex = 4
        Else
            Range("j" & rowIncrement).Interior.ColorIndex = 3
        End If
        Range("k" & rowIncrement).Value = ((Cells(a, 6).Value - stored_open) / stored_open) 'Write yearly change percentage value
        Range("k" & rowIncrement).NumberFormat = "0.00%"
        Range("l" & rowIncrement).Value = total_volume + Cells(a, 7).Value 'write total volume
        
        rowIncrement = rowIncrement + 1
        yearly_change = 0
        total_volume = 0
        newTickerRow = True
    Else
        If newTickerRow Then
            stored_open = Cells(a, 3).Value 'Store open value if it's the first row for a ticker symbol
            newTickerRow = False
        End If
        'yearly_change = yearly_change + (Cells(a, 6).Value - Cells(a, 3).Value)
        total_volume = total_volume + Cells(a, 7).Value
    End If
    
    
Next a

End Sub

