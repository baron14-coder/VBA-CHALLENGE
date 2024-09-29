# VBA-CHALLENGE
Sub multiple_year_stock_data()
Dim ws As Worksheet
For Each ws In Worksheets


Dim x As Long
Dim ticker As String
Dim quartely_change As Double
Dim total_stock_volume As Double



Dim title_row As Integer

Dim percentage_change As Double
percentage_change = 0



Dim coloumn As Integer
Dim lastrow As Long
title_row = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim quarterly_change_last_row As Long
quarterly_change_last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row
 
Dim last_percentage_change_row As Long
Dim last_total_stock_row As Long
last_percentage_change_row = ws.Cells(Rows.Count, 11).End(xlUp).Row
last_total_stock_row = ws.Cells(Rows.Count, 12).End(xlUp).Row



coloumn = 1
total_stock_volume = 0
quartely_change = 0
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Quarterly change"
ws.Range("K1") = "percentage change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"


For x = 2 To lastrow
If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then

ticker = ws.Cells(x, 1).Value
total_stock_volume = total_stock_volume + ws.Cells(x, 7).Value
quartely_change = ws.Cells(x, 6).Value - ws.Cells(x - 61, 3).Value
percentage_change = (quartely_change / ws.Cells(x - 61, 3).Value) * 100
ws.Range("L" & title_row).Value = total_stock_volume
ws.Range("I" & title_row).Value = ticker
ws.Range("J" & title_row).Value = quartely_change
ws.Range("K" & title_row).Value = percentage_change
title_row = title_row + 1
total_stock_volume = 0
quartely_change = 0
percentage_change = 0
Else
total_stock_volume = total_stock_volume + Cells(x, 7).Value
End If
Next x


Dim gpi As Double
gpi = Application.Max(ws.Range("k2:k" & last_percentage_change_row))
ws.Range("Q2").Value = gpi
Dim gpi_ticker As String

For x = 2 To last_percentage_change_row
If ws.Cells(x, 11).Value = gpi Then
gpi_ticker = ws.Cells(x, 9).Value
ws.Range("P2").Value = gpi_ticker
ws.Range("O2").Value = "Greatest % increase"
End If
Next x
Dim gpd As Double
gpd = Application.Min(ws.Range("K2:K" & last_percentage_change_row))
ws.Range("Q3").Value = gpd

Dim gpd_ticker As String
For x = 2 To last_percentage_change_row
If ws.Cells(x, 11).Value = gpd Then
gpd_ticker = ws.Cells(x, 9)
ws.Range("p3").Value = gpd_ticker
ws.Range("O3").Value = "Greatest % Decrease"

End If
Next x






Dim greatest_stock_volume As Double
greatest_stock_volume = Application.Max(ws.Range("L2:L" & last_total_stock_row))
ws.Range("Q4") = greatest_stock_volume
Dim grand_stock_name As String
For x = 2 To last_total_stock_row
If ws.Cells(x, 12).Value = greatest_stock_volume Then
grand_stock_name = ws.Cells(x, 9).Value
ws.Range("P4").Value = grand_stock_name
ws.Range("O4").Value = "Grand Total Volume"

End If
Next x
For x = 2 To quarterly_change_last_row
If ws.Cells(x, 11).Value > 0 Then
ws.Cells(x, 11).Interior.ColorIndex = 4
ws.Cells(x, 11).Font.ColorIndex = 1


ElseIf ws.Cells(x, 11).Value < 0 Then
ws.Cells(x, 11).Interior.ColorIndex = 3
ws.Cells(x, 11).Font.ColorIndex = 1
Else
ws.Cells(x, 11).Interior.ColorIndex = 2
ws.Cells(x, 11).Font.ColorIndex = 1
End If
Next x





Next ws



End Sub
