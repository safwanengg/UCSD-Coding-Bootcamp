Sub Stock()

For Each ws In Worksheets
' Set an initial variable for holding the ticker name
  Dim Ticker As String

  ' Set an initial variable for holding the total stock volume
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Keep track of the location for each ticker value in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Total Stock Volume"

  ' Loop through all ticker values
  For i = 2 To Last_Row

    ' Check if we are still within the same ticker value, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker value
      Ticker = ws.Cells(i, 1).Value

      ' Add to the total stock volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

      ' Print the ticker value in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the total stock volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total stock volume
      Total_Stock_Volume = 0

    ' If the cell immediately following a row is the same ticker value...
    Else

      ' Add to the total stock volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  
Next ws

End Sub
