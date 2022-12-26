VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SymbolTable As cSymbolTable

Private Sub Class_Initialize()

Set SymbolTable = New cSymbolTable
    initzSymbolTable SymbolTable
End Sub

Private Sub initzSymbolTable(ByRef inputSymbTab As cSymbolTable)
    With inputSymbTab
            .cAdd "SP", 0
            .cAdd "LCL", 1
            .cAdd "ARG", 2
            .cAdd "THIS", 3
            .cAdd "THAT", 4
            .cAdd "SCREEN", 16384
            .cAdd "KBD", 24576
            .cAdd "R0", 0
            .cAdd "R1", 1
            .cAdd "R2", 2
            .cAdd "R3", 3
            .cAdd "R4", 4
            .cAdd "R5", 5
            .cAdd "R6", 6
            .cAdd "R7", 7
            .cAdd "R8", 8
            .cAdd "R9", 9
            .cAdd "R10", 10
            .cAdd "R11", 11
            .cAdd "R12", 12
            .cAdd "R13", 13
            .cAdd "R14", 14
            .cAdd "R15", 15
    End With
End Sub

Public Function SymbolToBinStr(ByRef inputSymbolName As String) As String
    SymbolToBinStr = CustomFunctions.DecToBinStr(SymbolTable.cSearch(inputSymbolName))
End Function
