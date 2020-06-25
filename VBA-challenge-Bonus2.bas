Attribute VB_Name = "Module3"
Sub VBA_challenge_bonus2()
 
 For Each ws In Worksheets

    ' Create a summary table heading
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
      
      ' Set an initial variable for holding the ticket type name
      Dim Ticker_Name As String
    
      ' Set an initial variable for holding the total volume per ticker type
      Dim Volume_Total As Double
      Volume_Total = 0
      
      ' Set initial variables needed to calcuate Yearly and Percent Change and set the Open Value for first ticker in table
      Dim Open_Value, Close_Value, Yearly_Change As Double
      Open_Value = ws.Range("C2").Value
    
      ' Keep track of the location for each ticker type in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
    
      ' Identify the last row of the data
      Dim Last_Row As Long
      Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
      ' Loop through all data
      For i = 2 To Last_Row
    
        ' Check if we are still within the same ticker type, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Set the Ticker name
          Ticker_Name = ws.Cells(i, 1).Value
    
          ' Add to the Volume Total
          Volume_Total = Volume_Total + ws.Cells(i, 7).Value
          
          'Identify Close Value of Ticker
          Close_Value = ws.Cells(i, 6).Value
               
          'Calculate Yearly and Percent Change
          Yearly_Change = Close_Value - Open_Value
          If Open_Value <> 0 Then
            Percent_Change = Yearly_Change / Open_Value
          Else
            Percent_Change = "N/A"
          End If
          
          'Identify new Open Value for next Ticker (rewrite the previous one)
          Open_Value = ws.Cells(i + 1, 3).Value
               
          ' Print data to Summary Table:
          ' - Ticker:
          ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
          ' - Total Volume per ticker:
          ws.Range("M" & Summary_Table_Row).Value = Volume_Total
          ' - Yearly Change and setup the conditional formatting for positive and negative change:
          ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
            If Yearly_Change < 0 Then
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
          ' - Percent Change:
          ws.Range("L" & Summary_Table_Row).Value = Format(Percent_Change, "0.00%")
          
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Volume Total
          Volume_Total = 0
    
        ' If the cell immediately following a row is the same ticket type
        Else
    
          ' Add to the Volume Total
          Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    
        End If
    
      Next i
    
    'set the variables needed to calculate "Greatest % increase", "Greatest % decrease" and "Greatest total volume"'
    Dim max_percent, min_percent As Double
    Dim max_volume As Double
    Dim last_row_2 As Double
    
    'identify the last row of summary table as last_row_2
    last_row_2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'set the headings
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'calculate and write down "Greatest % increase"
    max_percent = WorksheetFunction.Max(ws.Range("L2:L" & last_row_2))
    max_ticker = WorksheetFunction.Match(max_percent, ws.Range("L2:L" & last_row_2), 0)
    ws.Cells(2, 16).Value = ws.Cells(max_ticker + 1, 10)
    ws.Cells(2, 17).Value = Format(max_percent, "0.00%")
    ws.Cells(2, 15).Value = "Greatest % increase"
    
    'calculate and write down "Greatest % decrease"
    min_percent = WorksheetFunction.Min(ws.Range("L2:L" & last_row_2))
    min_ticker = WorksheetFunction.Match(min_percent, ws.Range("L2:L" & last_row_2), 0)
    ws.Cells(3, 16).Value = ws.Cells(min_ticker + 1, 10)
    ws.Cells(3, 17).Value = Format(min_percent, "0.00%")
    ws.Cells(3, 15).Value = "Greatest % decrease"
    
    'calculate and write down "Greatest total volume"
    max_volume = WorksheetFunction.Max(ws.Range("M2:M" & last_row_2))
    max_ticker2 = WorksheetFunction.Match(max_volume, ws.Range("M2:M" & last_row_2), 0)
    ws.Cells(4, 16).Value = ws.Cells(max_ticker2 + 1, 10)
    ws.Cells(4, 17).Value = max_volume
    ws.Cells(4, 15).Value = "Greatest total volume"
  
  Next ws
  
End Sub

