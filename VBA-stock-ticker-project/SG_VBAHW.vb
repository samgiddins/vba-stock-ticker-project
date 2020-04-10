Sub Stonks():
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Open Price"
Cells(1, 11).Value = "Close Price"
Cells(1, 12).Value = "Yearly Change"
Cells(1, 13).Value = "Percent Change"
Cells(1, 14).Value = "Total Volume"

Dim Row As Long
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Volume As Double
    Total_Volume = 0
Dim Ticker_Index As Integer
Ticker_Index = 1

For Row = 2 To 800000
'Tickers
If Cells(Row, 1).Value <> Cells(Row - 1, 1).Value Then
Ticker_Index = Ticker_Index + 1
Cells(Ticker_Index, 9).Value = Cells(Row, 1).Value

'OpenPrice
Open_Price = Cells(Row, 3).Value
Cells(Ticker_Index, 10).Value = Open_Price
End If

'ClosePrice
If Cells(Row, 1).Value <> Cells(Row - 1, 1).Value And Row > 262 Then
Close_Price = Cells(Row - 1, 6).Value
Cells(Ticker_Index - 1, 11).Value = Close_Price
End If
Next Row

'Total Volume
Dim Volume_Row As Long
    Dim Volume_Index As Integer
        Volume_Index = 2
For Volume_Row = 2 To 800000
If Cells(Volume_Row + 1, 1).Value <> Cells(Volume_Row, 1).Value Then
Total_Volume = Total_Volume + Cells(Volume_Row, 7).Value
Volume_Index = Volume_Index + 1
Total_Volume = 0
Else
Total_Volume = Total_Volume + Cells(Volume_Row, 7).Value
End If
Cells(Volume_Index, 14).Value = Total_Volume
Next Volume_Row

'Yearly Change
Dim New_Row As Integer
For New_Row = 2 To 6000

If New_Row < 3170 Then
Yearly_Change = Cells(New_Row, 11).Value - Cells(New_Row, 10).Value
Cells(New_Row, 12).Value = Yearly_Change

'Percent Change
Percent_Change = Abs(Cells(New_Row, 12).Value / Cells(New_Row, 10).Value) * 100
Cells(New_Row, 13).Value = Percent_Change
End If

If Cells(New_Row, 12).Value < 0 Then
Cells(New_Row, 12).Interior.ColorIndex = 3
ElseIf Cells(New_Row, 12).Value > 0 Then
Cells(New_Row, 12).Interior.ColorIndex = 4
End If
Next New_Row

'Challenge
Cells(1, 17).Value = "Ticker"
Cells(1, 18).Value = "Value"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"

'Greatest % Increase
If Cells(New_Row, 13).Value = Application.WorksheetFunction.Max(Cells(New_Row, 13).Value) Then
Cells(2, 17).Value = Cells(Ticker_Index, 9).Value
Cells(2, 18).Value = Cells(New_Row, 12).Value
End If

'Greatest % Decrease
If Cells(New_Row, 13).Value = Application.WorksheetFunction.Min(Cells(New_Row, 13).Value) Then
Cells(3, 17).Value = Cells(Ticker_Index, 9).Value
Cells(3, 18).Value = Cells(New_Row, 12).Value
End If

'Greatest Total Volume
If Cells(Volume_Index, 14).Value = Application.WorksheetFunction.Max(Cells(Volume_Index, 14).Value) Then
Cells(4, 17).Value = Cells(Ticker_Index, 9).Value
Cells(4, 18).Value = Cells(New_Row, 12).Value
End If

End Sub