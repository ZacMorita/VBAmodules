Attribute VB_Name = "CustomFunctions"
Option Explicit

Public Function ConvertRangeToHTMLTable(rInput As Range) As String
    Dim rRow As Range
    Dim rCell As Range
    Dim strReturn As String
    Dim tdTag As String
    Dim trTag As String
    Dim RowIncrement As Long
    
    'Define table format and font
    strReturn = "<Table border='1' cellspacing='0' cellpadding='7' style='border-collapse:collapse;border:none'>  "
    tdTag = "<td valign='Center' style='border:solid windowtext 1.0pt; padding:0cm 5.4pt 0cm 5.4pt;height:1.05pt'>"
    trTag = " <tr align='Center'; style='height:10.00pt'> "
    
    'Loop through each row in the range
    For Each rRow In rInput.Rows
        RowIncrement = RowIncrement + 1
        'Start new html row
        strReturn = strReturn & trTag
        
        For Each rCell In rRow.Cells
            'If it is first or last row, then it is header or footer row, and will be bold
            If (RowIncrement = 1) Or RowIncrement = rInput.Rows.Count Then
                strReturn = strReturn & tdTag & "<b>" & rCell.Text & "</b></td>"
            Else
                strReturn = strReturn & tdTag & rCell.Text & "</td>"
            End If
        Next rCell
        'End a row
        strReturn = strReturn & "</tr>"
    Next rRow
    
    'Close the font tag
    strReturn = strReturn & "</font></table>"
    
    'Return html format
    ConvertRangeToHTMLTable = strReturn
End Function

Public Function FileExists(ByRef FullPath As String) As Boolean
    'Requires fully qualified path. Returns True if file (or folder) exists and False if it does not.
    Dim CheckDir As String
    
    'Dir returns the name of a Directory or File object if it exists or "" (vbNullString) if it finds nothing.
    CheckDir = Dir(FullPath, vbNormal)
    If Not CheckDir = vbNullString Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Function MoreThanAnHourAgo(ByRef MomentToTest As Date) As Boolean
    'Function returns true if input time is greater than an hour ago. _
        It returns false if input time is less than an hour ago _
        or if input time/date is greater that Now.
    If MomentToTest > Now Then
        MomentToTest = False
        Exit Function
    End If
    If DateValue(MomentToTest) = Date Then
        If Hour(Now - MomentToTest) > 0 Then
            MoreThanAnHourAgo = True
        End If
    Else
        MoreThanAnHourAgo = True
    End If
End Function
