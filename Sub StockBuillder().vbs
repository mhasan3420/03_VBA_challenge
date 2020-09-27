Sub StockBuillder()
Dim Yearly_open As Double
Dim Yearly_close As Double
Dim Yearly_change As Double
Dim Percentage_change As Double
Dim Total_volume As Double
Dim Summary_row As Double
Dim Row_Count As Long
Yearly_open = Cells(2, 3).Value
Total_volume = 0
Summary_row = 2
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Row_Count = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To Row_Count
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       Range("I" & Summary_row).Value = Cells(i, 1).Value
         'Yearly Change
      Yearly_close = Cells(i, 6).Value
      Yearly_change = (Yearly_close - Yearly_open)
      Range("J" & Summary_row).Value = Yearly_change
           'Conditional Foremating
        If Yearly_change > 0 Then
            Range("J" & Summary_row).Interior.ColorIndex = 4
        Else
            Range("J" & Summary_row).Interior.ColorIndex = 3
        End If
        If (Yearly_open = 0 And Yearly_close = 0) Then
                Percentage_change = 0
        ElseIf (Yearly_open = 0 And Yearly_close <> 0) Then
                Percentage_change = 1
        Else
              Percentage_change = (Yearly_change / Yearly_open)
              Range("K" & Summary_row).Value = Percentage_change
              Range("K" & Summary_row).NumberFormat = "0.00%"
        End If
                'Total Stock Volume
             Total_volume = Total_volume + Cells(i, 7).Value
             Range("L" & Summary_row).Value = Total_volume
             Summary_row = Summary_row + 1
             'Reset Stock Volume and Open Price
            Total_volume = 0
            Yearly_open = Cells(i + 1, 3).Value
    Else
        Total_volume = Total_volume + Cells(i, 7).Value
    End If
Next i
End Sub
