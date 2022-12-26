VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSymbolTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private mSymbolTable() As Variant
Private NextMemoryLoc As Long

Private Sub Class_Initialize()
    NextMemoryLoc = 16
    ReDim Preserve mSymbolTable(1, 0)
End Sub

Public Function cAdd(ByRef SymbolName As String, Optional ByRef SymbolValue = 0) As Long
    If Not IsEmpty(mSymbolTable(0, UBound(mSymbolTable, 2))) Then
        ReDim Preserve mSymbolTable(1, UBound(mSymbolTable, 2) + 1)
    End If
    mSymbolTable(0, UBound(mSymbolTable, 2)) = SymbolName
    mSymbolTable(1, UBound(mSymbolTable, 2)) = SymbolValue
    cAdd = mSymbolTable(1, UBound(mSymbolTable, 2))
End Function

Public Function cSearch(ByRef SymbolName As String) As Long
    Dim y As Long
    Dim found As Boolean: found = False

    For y = 0 To UBound(mSymbolTable, 2)
        If mSymbolTable(0, y) = SymbolName Then
            cSearch = mSymbolTable(1, y)
            found = True
            Exit For
        End If
    Next y
    
    If Not found Then
        cSearch = cAdd(SymbolName, NextMemoryLoc)
        NextMemoryLoc = NextMemoryLoc + 1
    End If
End Function
