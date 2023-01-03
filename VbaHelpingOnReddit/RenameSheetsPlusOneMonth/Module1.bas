Attribute VB_Name = "Module1"
Option Explicit

Private Sub Setup()
    Dim wb As Workbook
    Dim y As Integer
    Dim sd As Date
    
    Set wb = ThisWorkbook
    sd = "01/01/23"
    
    For y = 0 To 30
        With wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            .Name = Format(sd + y, "dd-mm-yy")
        End With
    Next y
End Sub

Private Sub TearDown()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Name = "Total Sum" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

Public Sub RenameSheetsPlusOneMonth()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        With ws
            If Not .Name = "Total Sum" Then
                .Name = AddOneMidString(.Name)
            End If
        End With
    Next ws
End Sub

Private Function AddOneMidString(ByVal InputDateString As String) As String
    Dim ReturnDateString As String
    Dim SplitUp() As String
    
    ReturnDateString = InputDateString
    SplitUp = Split(InputDateString, "-")

    If UBound(SplitUp) > 0 Then
        If IsNumeric(SplitUp(1)) Then
            If SplitUp(1) = 12 Then
                SplitUp(1) = "01"
                SplitUp(2) = Format(SplitUp(2) + 1, "00")
            Else
                SplitUp(1) = Format(SplitUp(1) + 1, "00")
            End If
        End If
        ReturnDateString = Join(SplitUp, "-")
    End If
    AddOneMidString = ReturnDateString
End Function
