VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Mnemonics As cMnemonics

Private Sub Class_Initialize()
    Set Mnemonics = New cMnemonics
End Sub

Public Function ConvertInstruction(ByRef inputAssemblyInst As String) As String
    Dim hString As String
    Dim Split1() As String
    Dim Split2() As String
    
    If InStr(1, inputAssemblyInst, "=", vbTextCompare) > 0 Then
        Split1 = Split(inputAssemblyInst, "=", -1, vbTextCompare)
        If InStr(1, Split1(1), ";", vbTextCompare) > 0 Then
            Split2 = Split(Split1(1), ";", -1, vbTextCompare)
            hString = "111" & _
              Mnemonics.GetComp(Split2(0)) & _
              Mnemonics.GetDest(Split1(0)) & _
              Mnemonics.GetJump(Split2(1))
        Else
            hString = "111" & _
              Mnemonics.GetComp(Split1(1)) & _
              Mnemonics.GetDest(Split1(0)) & _
              Mnemonics.GetJump("null")
        End If
    ElseIf InStr(1, inputAssemblyInst, ";", vbTextCompare) > 0 Then
        Split1 = Split(inputAssemblyInst, ";", -1, vbTextCompare)
        hString = "111" & _
            Mnemonics.GetComp(Split1(0)) & _
            Mnemonics.GetDest("null") & _
            Mnemonics.GetJump(Split1(1))
    Else
        hString = "111" & _
            Mnemonics.GetComp(inputAssemblyInst) & _
            Mnemonics.GetDest("null") & _
            Mnemonics.GetJump("null")
    End If
    
    ConvertInstruction = hString

End Function

