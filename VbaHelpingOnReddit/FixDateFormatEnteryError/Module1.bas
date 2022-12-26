Attribute VB_Name = "Module1"
Option Explicit

Public Function FixBadDateFormatting(ByRef InputString As String) As String
    'This function can also be used a a formula directly from a worksheet _
        by typing something like "=FixBadDateFormatting(T17)" into the formula bar
    Dim AfterReplace As String

    'replace does nothing to the string if it doesn't find anything
    AfterReplace = Replace(InputString, ".", "/")

    'see if it's a date after replace
    If IsDate(AfterReplace) Then
        'change to date
        FixBadDateFormatting = DateValue(AfterReplace)
    Else
        'return original string if not a date after replace
        FixBadDateFormatting = InputString
    End If
End Function

Public Sub ChangeEachStringToDateInRange()
    Dim CheckingRange As Range
    Dim CurrentCell As Range
    
    'set the pointer as a range in the currently active sheet
    Set CheckingRange = ActiveSheet.Range("$A$7:$A$28")
    'optionally you could set CheckingRange to "Selection" to only _
        check ranges you have selected.

    For Each CurrentCell In CheckingRange.Cells
        'for each cell in range, use the FixBadDateFormMatting function.
        CurrentCell.Value = FixBadDateFormatting(CurrentCell.Value)
    Next CurrentCell
End Sub
