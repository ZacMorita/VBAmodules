Attribute VB_Name = "Module1"
Option Explicit

Private Function CountColumnsWithColoredCell(ByRef rng As Range, _
    Optional ByRef ColorToCount As Variant = 38) As Long
    'NOTE: DisplayFormat is the product of conditional formatting. _
        This function will return #VALUE! if called by worksheet _
        formula. To use, call from VBA Sub.
    Dim CurrentCol As Range
    Dim CurrentCel As Range
    Dim ColorCounter As Long

    'Go through each column
    For Each CurrentCol In rng.Columns
        'go through each cell of current column
        For Each CurrentCel In CurrentCol.Cells
            'ColorIndex 38 is the same as RGB(255, 199, 206) [Light Red]
            If CurrentCel.DisplayFormat.Interior.ColorIndex = ColorToCount Then
                ColorCounter = ColorCounter + 1 'counts column
                Exit For 'stops and goes to next column
            End If
        Next CurrentCel
    Next CurrentCol
    CountColumnsWithColoredCell = ColorCounter
End Function
'--------------------------------------
Public Sub UseCountFunctOnSelection()
    MsgBox "The total columns in current selection containing at least " & _
        "one colored cell is: " & _
        CountColumnsWithColoredCell(Selection)
End Sub
'--------------------------------------
Public Sub UseCountFunctForEachName()
    Dim wb As Workbook
    Dim CurrentName As Name
    Dim total As Long
    
    Set wb = ThisWorkbook
    
    For Each CurrentName In wb.Names
        total = total + CountColumnsWithColoredCell(CurrentName.RefersToRange)
    Next CurrentName
    MsgBox "The total columns in all named ranges in this workbook " & _
           "containing at least one colored cell is: " & total
End Sub
