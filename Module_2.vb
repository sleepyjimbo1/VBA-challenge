Sub Module_2()

'Loop through each sheet
For Each ws In Worksheets

'Declare the variables
Dim Ticker As String
Dim Row As Integer

Dim greatest_ticker As String
Dim greatest_increase As Double
Dim lowest_ticker As String
Dim greatest_decrease As Double
Dim total_ticker As String
Dim greatest_total As Double

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Input the Header Text
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

Row = 2

'Testing a smaller sample
For I = 2 To LastRow

close_price = ws.Cells(I, 6).Value
open_price = ws.Cells(I, 3).Value
volume = ws.Cells(I, 7).Value

'If the next row's ticker does not equal the current row, then input the previous ticker
If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then

x = close_price - open_price
Z = open_price
a = volume

yearly_change = yearly_change + x
denom = denom + Z
sum_volume = sum_volume + a

percent_change = (yearly_change / denom) * 100

ws.Cells(Row, 9).Value = Ticker
ws.Cells(Row, 10).Value = yearly_change
ws.Cells(Row, 11).Value = percent_change
ws.Cells(Row, 12).Value = sum_volume

Row = Row + 1

yearly_change = 0
denom = 0
percent_change = 0
sum_volume = 0

Else
Ticker = ws.Cells(I, 1).Value
x = close_price - open_price
yearly_change = yearly_change + x
a = volume
sum_volume = sum_volume + a

Z = open_price
denom = denom + Z

End If
Next I

'Greatest Percent Increase
greatest_increase = 0
For e = 2 To LastRow
    If ws.Cells(e, 11).Value > greatest_increase Then
        greatest_ticker = ws.Cells(e, 9).Value
        greatest_increase = ws.Cells(e, 11).Value
    End If
Next e

ws.Cells(2, 16).Value = greatest_increase
ws.Cells(2, 15).Value = greatest_ticker

'Greatest Percent Decrease
greatest_decrease = 0
For ee = 2 To LastRow
    If ws.Cells(ee, 11).Value < greatest_decrease Then
        lowest_ticker = ws.Cells(ee, 9).Value
        greatest_decrease = ws.Cells(ee, 11).Value
    End If
Next ee

ws.Cells(3, 16).Value = greatest_decrease
ws.Cells(3, 15).Value = lowest_ticker

'Greatest Total Volume
greatest_total = 0
For eee = 2 To LastRow
    If ws.Cells(eee, 12).Value > greatest_total Then
        total_ticker = ws.Cells(eee, 9).Value
        greatest_total = ws.Cells(eee, 12).Value
    End If
Next eee

ws.Cells(4, 16).Value = greatest_total
ws.Cells(4, 15).Value = total_ticker

'Color Coding Yearly Change
For e4 = 2 To LastRow
If ws.Cells(e4, 10).Value > 0 Then
    ws.Cells(e4, 10).Interior.Color = vbRed
    Else
    ws.Cells(e4, 10).Interior.Color = vbGreen
    End If
    Next e4

Next ws

    Sheets("2018").Select
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Sheets("2019").Select
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Sheets("2020").Select
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"

End Sub