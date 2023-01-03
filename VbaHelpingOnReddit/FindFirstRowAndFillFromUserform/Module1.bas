Attribute VB_Name = "Module1"
Option Explicit

#Const SubVersion = 2 '(1 = SpecialCellsVersion / 2 = FindVersion)

Private Sub Add_Click()
    #If SubVersion = 1 Then
        SpecialCellsVersion
    #ElseIf SubVersion = 2 Then
        FindVersion
    #End If
End Sub

#If SubVersion = 1 Then

    Private Sub SpecialCellsVersion()
        Dim lo As ListObject
        Dim spcrng As Range
        
        'Change Sheets(1) to Ark1
        Set lo = ThisWorkbook.Sheets(1).ListObjects(1)
        
        On Error Resume Next 'In case there are no blanks
        Set spcrng = lo.ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeBlanks).Cells(1)
        
        If Err.Number <> 0 And Err.Description = "No cells were found." Then
            'Adds a new row if no blanks were found
            Set spcrng = lo.ListRows.Add.Range.Cells(1, 1)
            Err.Clear 'Clears the expected/intentional error
        ElseIf Err.Number <> 0 Then
            On Error GoTo 0 'Allow default error popup to happen
            Err.Raise Err.Number 'Raise error popup
        End If
        
        With spcrng
            .Offset(0, 2).Value = 1 'Me.Box1.Value
            .Offset(0, 4).Value = 2 'Me.Box2.Value
            .Offset(0, 15).Value = 3 'Me.Box3.Value
            .Offset(0, 8).Value = 4 'Me.Box4.Value
            .Offset(0, 10).Value = 5 'Me.Box5.Value
            .Offset(0, 11).Value = 6 'Me.Box6.Value
            .Offset(0, 7).Value = 7 'Me.Box7.Value
            .Offset(0, 13).Value = 8 'Me.Text1.Value
        End With
    End Sub

#ElseIf SubVersion = 2 Then

    Private Sub FindVersion()
        Dim lo As ListObject
        Dim FindRng As Range
        Dim FoundRng As Range
        
        Set lo = ThisWorkbook.Sheets(1).ListObjects(1)
        Set FindRng = lo.ListColumns(3).Range
        Set FoundRng = FindRng.Find("", LookIn:=xlValues)
        
        If FoundRng Is Nothing Then
            Set FoundRng = lo.ListRows.Add.Range.Cells(1, 3)
        End If
        
        With FoundRng
            .Value = 1 'Me.Box1.Value
            .Offset(0, 2).Value = 2 'Me.Box2.Value
            .Offset(0, 13).Value = 3 'Me.Box3.Value
            .Offset(0, 6).Value = 4 'Me.Box4.Value
            .Offset(0, 8).Value = 5 'Me.Box5.Value
            .Offset(0, 9).Value = 6 'Me.Box6.Value
            .Offset(0, 5).Value = 7 'Me.Box7.Value
            .Offset(0, 11).Value = 8 'Me.Text1.Value
        End With
        
    End Sub

#End If
