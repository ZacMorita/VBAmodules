Attribute VB_Name = "Module1"
Option Explicit

Sub DoDates()
    Dim ws As Worksheet
    Dim DateRow As Long
    Dim DateCol As Long
    Dim i As Long
    Dim DateVar As Variant

    Set ws = ActiveSheet
    DateCol = 2 '<-- Change to the Column (number) the dates are in.
    DateRow = ws.Cells(1, DateCol).End(xlUp).Row
    DateVar = ws.Cells(DateRow, DateCol).Value

    If IsDate(DateVar) Then
        While (DateVar + i) < Date
            i = i + 1
            ws.Cells(DateRow + i, DateCol).Value = DateVar + i
        Wend
    Else
        MsgBox "The value: [" & DateVar & "] is not a date." & vbCrLf & _
               "The Address is : " & ws.Cells(DateRow, DateCol).Address
    End If

End Sub
