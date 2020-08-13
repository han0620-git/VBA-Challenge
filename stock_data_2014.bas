Attribute VB_Name = "Module1"
Sub stock_data()

  Dim ticker As String
  Dim open_price As Double
  Dim close_price As Double
  Dim yearly_change As Double
  Dim percent_change As Double
  
  Dim Stock_Total As Double
  Stock_Total = 0

  
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

  open_price = Cells(2, 3).Value
  
  If open_price = 0 Then
  open_price = 1
  
  End If
  
  
  For i = 2 To 70926

    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      
      ticker = Cells(i, 1).Value
      
      close_price = Cells(i, 6).Value
      
      yearly_change = close_price - open_price
      Range("J" & Summary_Table_Row).Value = yearly_change
      
      
      percent_change = 100 * (yearly_change / open_price)
      Range("K" & Summary_Table_Row).Value = percent_change
      
      open_price = Cells(i + 1, 3).Value
      
      Stock_Total = Stock_Total + Cells(i, 7).Value

      
      Range("I" & Summary_Table_Row).Value = ticker

      Range("L" & Summary_Table_Row).Value = Stock_Total
      
      
      Cells(1, 9).Value = "Ticker"
      Cells(1, 10) = "Yearly Change"
      Cells(1, 11) = "Percent Change"
      Cells(1, 12).Value = "Stock Total"

      
      Summary_Table_Row = Summary_Table_Row + 1
      
      
      Stock_Total = 0

    
    Else

      Stock_Total = Stock_Total + Cells(i, 7).Value
        
    
    End If

  Next i
    
      
End Sub
