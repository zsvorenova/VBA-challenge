Attribute VB_Name = "Module2"
Sub VBA_challenge_bonus1()

'set the variables needed to calculate "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
Dim max_percent, min_percent As Double
Dim max_volume As Double
Dim last_row_2 As Double

'identify the last row of summary table as last_row_2
last_row_2 = Cells(Rows.Count, 10).End(xlUp).Row

'set the headings
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'calculate and write down "Greatest % increase"
max_percent = WorksheetFunction.Max(Range("L2:L" & last_row_2))
max_ticker = WorksheetFunction.Match(max_percent, Range("L2:L" & last_row_2), 0)
Cells(2, 16).Value = Cells(max_ticker + 1, 10)
Cells(2, 17).Value = Format(max_percent, "0.00%")
Cells(2, 15).Value = "Greatest % increase"

'calculate and write down "Greatest % decrease"
min_percent = WorksheetFunction.Min(Range("L2:L" & last_row_2))
min_ticker = WorksheetFunction.Match(min_percent, Range("L2:L" & last_row_2), 0)
Cells(3, 16).Value = Cells(min_ticker + 1, 10)
Cells(3, 17).Value = Format(min_percent, "0.00%")
Cells(3, 15).Value = "Greatest % decrease"

'calculate and write down "Greatest total volume"
max_volume = WorksheetFunction.Max(Range("M2:M" & last_row_2))
max_ticker2 = WorksheetFunction.Match(max_volume, Range("M2:M" & last_row_2), 0)
Cells(4, 16).Value = Cells(max_ticker2 + 1, 10)
Cells(4, 17).Value = max_volume
Cells(4, 15).Value = "Greatest total volume"

End Sub


