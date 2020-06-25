Attribute VB_Name = "Module1"
Sub VBA_challenge_HW()

  ' Create a summary table heading
  Range("J1").Value = "Ticker"
  Range("K1").Value = "Yearly Change"
  Range("L1").Value = "Percent Change"
  Range("M1").Value = "Total Stock Volume"
      
  ' Set an initial variable for holding the ticket type name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total volume per ticker type
  Dim Volume_Total As Double
  Volume_Total = 0
  
  ' Set initial variables needed to calcuate Yearly and Percent Change and set the Open Value for first ticker in table
  Dim Open_Value, Close_Value, Yearly_Change As Double
  Open_Value = Range("C2").Value

  ' Keep track of the location for each ticker type in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Identify the last row of the data
  Dim Last_Row As Long
  Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
   
  ' Loop through all data
  For i = 2 To Last_Row

    ' Check if we are still within the same ticker type, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
      
      'Identify Close Value of Ticker
      Close_Value = Cells(i, 6).Value
           
      'Calculate Yearly and Percent Change
      Yearly_Change = Close_Value - Open_Value
      If Open_Value <> 0 Then
        Percent_Change = Yearly_Change / Open_Value
      Else
        Percent_Change = "N/A"
      End If
      
      'Identify new Open Value for next Ticker (rewrite the previous one)
      Open_Value = Cells(i + 1, 3).Value
           
      ' Print data to Summary Table:
      ' - Ticker:
      Range("J" & Summary_Table_Row).Value = Ticker_Name
      ' - Total Volume per ticker:
      Range("M" & Summary_Table_Row).Value = Volume_Total
      ' - Yearly Change and setup the conditional formatting for positive and negative change:
      Range("K" & Summary_Table_Row).Value = Yearly_Change
        If Yearly_Change < 0 Then
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        Else
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        End If
      ' - Percent Change:
      Range("L" & Summary_Table_Row).Value = Format(Percent_Change, "0.00%")
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total
      Volume_Total = 0

    ' If the cell immediately following a row is the same ticket type
    Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If

  Next i

End Sub

