Attribute VB_Name = "Main"
Option Explicit
Option Base 0

Sub RunAssembler()
    Dim Picker As FileDialog
    Dim Parser As cParser
    Dim ReadPath As String
    Dim WritePath As String
    
    Set Parser = New cParser
    Set Picker = Application.FileDialog(msoFileDialogFilePicker)
    
    With Picker
        If .Show = -1 Then
            ReadPath = .SelectedItems(1)
        End If
    End With
    
    WritePath = Replace(ReadPath, ".asm", ".hack", 1, , vbTextCompare)
    
    Parser.RunReadWrite ReadPath, WritePath

End Sub
