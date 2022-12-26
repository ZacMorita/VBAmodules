Attribute VB_Name = "Module1"
Option Explicit

Public Sub SkipFilterdRows()
    'This module is intentionally generic for educational purposes. _
        It is meant to be direct. Not an example of preferred methods.
    Dim wb As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim rng As Range
    Dim cel As Range
    Dim MaxRows As Long
    Dim i As Long
    
    Set wb = ThisWorkbook
    Set ws1 = wb.Sheets(1)
    Set ws2 = wb.Sheets(2)
    Set rng = ws1.Range(ws1.Cells(2, 1), ws1.Cells(ws1.Rows.Count, 1).End(xlUp))
    i = 10 '<--- This assignment is the starting row of the copy-to range.
    MaxRows = ws2.Range("$G$2").Value + i 'max rows + starting row = final paste-to row
    
    For Each cel In rng
        If Not cel.RowHeight = 0 Then
            ws2.Cells(i, 1).Value = cel.Value
            ws2.Cells(i, 2).Value = cel.Offset(0, 1).Value
            i = i + 1
            If i = (MaxRows) Then
                'If ws2.Range("$G$2") = 0 then this sub will move all visible rows.
                Exit For
            End If
        End If
    Next cel
End Sub
