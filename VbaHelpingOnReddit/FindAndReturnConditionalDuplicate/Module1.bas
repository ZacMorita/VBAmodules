Attribute VB_Name = "Module1"
Option Explicit

Sub FindTheDuplicate()
    Dim LookInRng As Range
    Dim CurrentLookRng As Range
    Dim FoundRng As Range

    'set range to check
    Set LookInRng = ActiveSheet.Range("$A$1:$A$100")

    'go through range top down
    For Each CurrentLookRng In LookInRng
        'skip cells that contain "Agilent 5110"
        If Not InStr(1, CurrentLookRng.Value, "Agilent 5110") > 0 Then
            'Use find with first 11 characters of current cell as parameter
            Set FoundRng = LookInRng.Find(Left(CurrentLookRng.Value, 11), CurrentLookRng, xlValues, xlPart, xlNext)
            'if something was found
            If Not FoundRng Is Nothing Then
                'and it's not the same as current cell
                If Not FoundRng.Address = CurrentLookRng.Address Then
                    'then it is a duplicate and is placed in P3
                    ActiveSheet.Range("$P$3").Value = FoundRng.Value
                    ActiveSheet.Range("$P$2").Value = "QC" & (CInt(Mid(FoundRng.Value, 11, 1)) - 1)
                    'and the sub ends
                    Exit Sub
                End If
            End If
        End If
    Next CurrentLookRng
End Sub
'==========================
'vvvvv Results VVVVV
'==========================
' Range("P2") = "QC7"
' Range("P3") = "M1111111118something10"
'==========================
'VVVVV Test Data VVVVV
'==========================
'COLUMN A
'--------
'M1111111111something1
'M1111111112something2
'M1111111113something3
'Agilent 5110something4
'M1111111115something5
'M1111111116something6
'M1111111117something7
'M1111111118something8
'M1111111119something9
'M1111111118something10
'M1111111121something11
'M1111111122something12
'Agilent 5110something13
'M1111111124something14
'M1111111125something15
'M1111111118something16
'M1111111127something17
'M1111111128something18
'M1111111129something19
'M1111111130something20
'M1111111118something21
'M1111111132something22
'M1111111133something23
'M1111111118something24
'M1111111118something25

