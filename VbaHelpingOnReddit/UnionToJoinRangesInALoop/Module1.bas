Attribute VB_Name = "Module1"
Option Explicit

Sub SelectAreasThatAreIndented()
    Dim rng As Range
    Dim cel As Range
    Dim hold As Range
    'Application.Union prevents multiple selections of the same range, _
        so this works on 2d ranges as well.
        
    'Change this range to the one you're checking in
    Set rng = Range("$A$1:$B$32")

    For Each cel In rng
        'Check for leading space (chr(32))
        If cel.IndentLevel > 0 Then
            'If first match then set, else join with previous found
            If hold Is Nothing Then
                Set hold = cel.EntireRow
            Else
                Set hold = Application.Union(hold, cel.EntireRow)
            End If
        End If
    Next cel
    hold.Select 'Selects all the areas
End Sub
'-----------------------------------------
Sub SelectAreasStartingWithSpaces()
    Dim rng As Range
    Dim cel As Range
    Dim hold As Range
    'Application.Union prevents multiple selections of the same range, _
        so this works on 2d ranges as well.
        
    'Change this range to the one you're checking in
    Set rng = Range("$A$1:$B$32")

    For Each cel In rng
        'Check for leading space (chr(32))
        If Left(cel.Value, 1) = " " Then
            'If first match then set, else join with previous found
            If hold Is Nothing Then
                Set hold = cel.EntireRow
            Else
                Set hold = Application.Union(hold, cel.EntireRow)
            End If
        End If
    Next cel
    hold.Select 'Selects all the areas
End Sub
