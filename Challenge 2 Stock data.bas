Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

Dim i As Integer

Dim ws As Worksheet

For Each ws In Worksheets

Dim LastRow As Long

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("j1").EntireColumn.Insert
ws.Cells(1, 10).Value = "Ticker"
ws.Range("k1").EntireColumn.Insert
ws.Cells(1, 11).Value = "Quarterly Price Change"
ws.Range("l1").EntireColumn.Insert
ws.Cells(1, 12).Value = "Percent Change %"
ws.Range("m1").EntireColumn.Insert
ws.Cells(1, 13).Value = "Total Stock Volume"

Dim Ticker As String
Dim Open_Price As Double
Dim Closing_Price As Double
Dim Quarterly_Price_Change As Double
Dim Percent_Change As Double


Dim Stock_Volume As Double
Stock_Volume = 0

Dim Summary_Table As Integer
Summary_Table = 2

Open_Price = ws.Cells(2, 3).Value

For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Ticker = ws.Cells(i, 1).Value
Closing_Price = ws.Cells(i, 6).Value
Quarterly_Price_Change = Closing_Price - Open_Price
Stock_Volume = Stock_Volume + Cells(i, 7).Value
Percent_Change = Quarterly_Price_Change / Open_Price


ws.Range("J" & Summary_Row_Table).Value = Ticker
ws.Range("K" & Summary_Row_Table).Value = Quarterly_Price_Change
ws.Range("L" & Summary_Row_Table).Value = Percent_Change
ws.Range("M" & Summary_Row_Table).Value = Stock_Volume

Summary_Row_Table = Summay_Row_Table + 1

Stock_Volume = 0

Open_Price = ws.Cells(i + 1, 3).Value

Else

Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value


End If

Next i


Next ws


End Sub






