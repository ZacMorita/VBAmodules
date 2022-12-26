VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMnemonics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CompMnemonics As Collection
Private DestMnemonics As Collection
Private JumpMnemonics As Collection

Private Sub Class_Initialize()
    Set CompMnemonics = initzCompMnem(CompMnemonics)
    Set DestMnemonics = initzDestMnem(DestMnemonics)
    Set JumpMnemonics = initzJumpMnem(JumpMnemonics)
End Sub

Private Function initzCompMnem(ByRef inputCollection As Collection) As Collection
    Set inputCollection = New Collection
        With inputCollection
            .Add "0101010", "0"
            .Add "0111111", "1"
            .Add "0111010", "-1"
            .Add "0001100", "D"
            .Add "0110000", "A"
            .Add "0001101", "!D"
            .Add "0110001", "!A"
            .Add "0001111", "-D"
            .Add "0110011", "-A"
            .Add "0011111", "D+1"
            .Add "0110111", "A+1"
            .Add "0001110", "D-1"
            .Add "0110010", "A-1"
            .Add "0000010", "D+A"
            .Add "0010011", "D-A"
            .Add "0000111", "A-D"
            .Add "0000000", "A&D"
            .Add "0010101", "D|A"
            .Add "1110000", "M"
            .Add "1110001", "!M"
            .Add "1110011", "-M"
            .Add "1110111", "M+1"
            .Add "1110010", "M-1"
            .Add "1000010", "D+M"
            .Add "1010011", "D-M"
            .Add "1000111", "M-D"
            .Add "1000000", "D&M"
            .Add "1010101", "D|M"
        End With
    Set initzCompMnem = inputCollection
End Function

Private Function initzDestMnem(ByRef inputCollection As Collection) As Collection
    Set inputCollection = New Collection
        With inputCollection
            .Add "000", "null"
            .Add "001", "M"
            .Add "010", "D"
            .Add "011", "MD"
            .Add "100", "A"
            .Add "101", "AM"
            .Add "110", "AD"
            .Add "111", "AMD"
        End With
    Set initzDestMnem = inputCollection
End Function

Private Function initzJumpMnem(ByRef inputCollection As Collection) As Collection
    Set inputCollection = New Collection
        With inputCollection
            .Add "000", "null"
            .Add "001", "JGT"
            .Add "010", "JEQ"
            .Add "011", "JGE"
            .Add "100", "JLT"
            .Add "101", "JNE"
            .Add "110", "JLE"
            .Add "111", "JMP"
        End With
    Set initzJumpMnem = inputCollection
End Function

Public Function GetComp(ByRef inputString As String) As String
    GetComp = CompMnemonics.Item(inputString)
End Function

Public Function GetDest(ByRef inputString As String) As String
    GetDest = DestMnemonics.Item(inputString)
End Function

Public Function GetJump(ByRef inputString As String) As String
    GetJump = JumpMnemonics.Item(inputString)
End Function
