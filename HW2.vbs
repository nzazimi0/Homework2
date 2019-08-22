Attribute VB_Name = "Module1"
Sub volumess()

  ' Set an initial variable for holding the brand name
  Dim ticker As String

  ' Set an initial variable for holding the total per credit card brand
  Dim volume As Double
  volume = 0

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
  ' Loop through all credit card purchases
  For i = 2 To RowCount

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        

      ' Set the Brand name
        ticker = Cells(i, 1).Value

      ' Add to the Brand Total
      volume = volume + Cells(i, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("I" & Summary_Table_Row).Value = Cells(i, 1).Value

      ' Print the Brand Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
    volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      volume = volume + Cells(i, 7).Value

    End If

  Next i

End Sub



