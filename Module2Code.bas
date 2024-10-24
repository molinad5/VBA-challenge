Attribute VB_Name = "Module2"
Option Explicit

Sub reset()

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
   
    ws.Activate

ws.Range("I:ZZ").Delete

Worksheets("Q1").Select

Next ws

End Sub
