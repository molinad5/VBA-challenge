Attribute VB_Name = "Module1"
Option Explicit

Sub tickerStock()

    ' Worksheet loop
    Dim ws As Worksheet
    Dim last_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim quarterly_change As Double
    Dim ticker As String
    Dim percent_change As Double
    Dim volume As Double
    Dim row As Long
    Dim column As Integer
    Dim i As Long, j As Long, k As Long
    Dim quarterly_change_last_row As Long

    ' Loop through each worksheet
    For Each ws In ActiveWorkbook.Worksheets
        
        ' Find the last row of the table
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        ' Add headers if not already present
        If ws.Cells(1, 9).Value = "" Then
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
        End If

        volume = 0
        row = 2
        column = 1
        
        ' Set the initial price
        open_price = ws.Cells(2, column + 2).Value
        
        ' Loop through all tickers to check for changes
        For i = 2 To last_row
            
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                
                ' Set ticker name
                ticker = ws.Cells(i, column).Value
                ws.Cells(row, column + 8).Value = ticker
                
                ' Set close price
                close_price = ws.Cells(i, column + 5).Value
                
                ' Calculate quarterly change
                quarterly_change = close_price - open_price
                ws.Cells(row, column + 9).Value = quarterly_change
                
                ' Calculate percent change
                If open_price <> 0 Then
                    percent_change = quarterly_change / open_price
                Else
                    percent_change = 0
                End If
                ws.Cells(row, column + 10).Value = percent_change
                ws.Cells(row, column + 10).NumberFormat = "0.00%"
                
                ' Calculate total volume per quarter
                volume = volume + ws.Cells(i, column + 6).Value
                ws.Cells(row, column + 11).Value = volume
                
                ' Move to the next row
                row = row + 1
                
                ' Reset open price for next ticker
                open_price = ws.Cells(i + 1, column + 2).Value
                
                ' Reset volume for next ticker
                volume = 0
                
            Else
                volume = volume + ws.Cells(i, column + 6).Value
            End If
        Next i
        
        ' Find the last row of the ticker column
        quarterly_change_last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
        
        ' Set cell colors based on changes
        For j = 2 To quarterly_change_last_row
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10  ' Green for positive or zero
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3   ' Red for negative
            End If
        Next j
        
        ' Set headers for greatest increases/decreases
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Find highest values for each ticker
        For k = 2 To quarterly_change_last_row
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quarterly_change_last_row)) Then
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quarterly_change_last_row)) Then
                ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quarterly_change_last_row)) Then
                ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
            End If
        Next k
        
        ' Format output
        ws.Range("I:Q").Font.Bold = True
        ws.Range("I:Q").EntireColumn.AutoFit
        
    Next ws

End Sub

