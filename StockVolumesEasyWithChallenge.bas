Attribute VB_Name = "Module1"
Sub challengeStockTablesWS()

  Dim ws As Worksheet
  For Each ws In Worksheets
  
  Dim Stock_Symbol As String
  Dim Stock_Total As Double
  Stock_Total = 0
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  
  For I = 2 To lastrow
    
 
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
      Stock_Symbol = ws.Cells(I, 1).Value
      Stock_Total = Stock_Total + ws.Cells(I, 7).Value
      ws.Range("I" & Summary_Table_Row).Value = Stock_Symbol
      ws.Range("J" & Summary_Table_Row).Value = Stock_Total


      Summary_Table_Row = Summary_Table_Row + 1
      Stock_Total = 0


    Else
      Stock_Total = Stock_Total + ws.Cells(I, 7).Value

    End If
  
  Next I
  
Next ws

End Sub


