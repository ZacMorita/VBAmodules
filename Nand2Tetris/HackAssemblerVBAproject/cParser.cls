VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum StringDecision
    SkpChar = 1
    SkpLine = 2
    Comment = 3
    AInstc = 4
    CInstc = 5
End Enum

Private Parser As FileSystemObject
Private coder As FileSystemObject
Private readtext As TextStream
Private writetext As TextStream

Public Instructions As cInstructions
Public Symbols As cSymbols

Private Sub Class_Initialize()
    Set Parser = New FileSystemObject
    Set coder = New FileSystemObject
    Set Instructions = New cInstructions
    Set Symbols = New cSymbols
End Sub

Public Sub RunReadWrite(ByRef ReadFrom As String, ByRef WriteTo As String)
    Set readtext = Parser.OpenTextFile(ReadFrom, ForReading, False, TristateUseDefault)
    
    FirstPass readtext
    readtext.Close
    
    Set readtext = Parser.OpenTextFile(ReadFrom, ForReading, False, TristateUseDefault)
    Set writetext = coder.OpenTextFile(WriteTo, ForWriting, True, TristateUseDefault)
    
    SecondPass readtext, writetext
End Sub

Private Sub FirstPass(ByRef inputTextStream As TextStream)
    Dim CurrentRead As String
    Dim InstructionCount As Long
    Dim NewSymbol As String
    Dim CurrentStringDecision As StringDecision
    
    NewSymbol = vbNullString
    InstructionCount = 0
    
    Do Until readtext.AtEndOfStream = True
        CurrentRead = readtext.Read(1)
        If CurrentRead = Chr(40) Then
            CurrentRead = readtext.Read(1)
            Do
                If inputTextStream.AtEndOfLine = True Then
                    NewSymbol = vbNullString
                    Exit Do
                Else
                    NewSymbol = NewSymbol & CurrentRead
                End If
                CurrentRead = readtext.Read(1)
            Loop Until CurrentRead = Chr(41)
            If Not NewSymbol = vbNullString Then
                Symbols.SymbolTable.cAdd NewSymbol, InstructionCount
                NewSymbol = vbNullString
                readtext.SkipLine
            End If
        Else
            CurrentStringDecision = ProcessString(CurrentRead)
            If CurrentStringDecision = Comment Then
                CurrentRead = readtext.Read(1)
                If CurrentRead = Chr(47) Then
                    readtext.SkipLine
                End If
            ElseIf CurrentStringDecision >= 4 Then
                InstructionCount = InstructionCount + 1
                readtext.SkipLine
            End If
        End If
    Loop

End Sub

Private Sub SecondPass(ByRef inputReadStream As TextStream, ByRef inputWriteStream As TextStream)
    Dim CurrentRead As String
    Dim NewInstruction As String
    Dim CurrentStringDecision As StringDecision
    
    NewInstruction = vbNullString
    
    Do Until readtext.AtEndOfStream = True
        CurrentRead = readtext.Read(1)
        CurrentStringDecision = ProcessString(CurrentRead)
        If CurrentStringDecision = Comment Then
            CurrentRead = readtext.Read(1)
            If CurrentRead = Chr(47) Then
                readtext.SkipLine
            End If
        ElseIf CurrentStringDecision = SkpLine Then
            readtext.SkipLine
        ElseIf CurrentStringDecision = AInstc Then
            CurrentRead = readtext.Read(1)
            Do
                NewInstruction = NewInstruction & CurrentRead
                CurrentRead = readtext.Read(1)
                CurrentStringDecision = ProcessString(CurrentRead)
            Loop Until CurrentStringDecision <= 3
            If IsNumeric(NewInstruction) Then
                writetext.WriteLine CustomFunctions.DecToBinStr(CInt(NewInstruction))
            Else
                writetext.WriteLine Symbols.SymbolToBinStr(NewInstruction)
            End If
        ElseIf CurrentStringDecision = CInstc Then
            Do
                NewInstruction = NewInstruction & CurrentRead
                CurrentRead = readtext.Read(1)
                CurrentStringDecision = ProcessString(CurrentRead)
            Loop Until CurrentStringDecision <= 3
            writetext.WriteLine Instructions.ConvertInstruction(NewInstruction)
        End If
        NewInstruction = vbNullString
    Loop
End Sub

Private Function ProcessString(ByRef inputString As String) As StringDecision
    Select Case inputString
        Case Is = Chr(32)
            ProcessString = SkpChar
        Case Is = Chr(10)
            ProcessString = SkpChar
        Case Is = Chr(13)
            ProcessString = SkpChar
        Case Is = Chr(9)
            ProcessString = SkpChar
        Case Is = Chr(40)
            ProcessString = SkpLine
        Case Is = Chr(47)
            ProcessString = Comment
        Case Is = Chr(64)
            ProcessString = AInstc
        Case Else
            ProcessString = CInstc
    End Select
End Function

Private Sub AssembleFromText(ByRef inputReadStream As TextStream, ByRef inputWriteStream As TextStream)
    Dim CurrentRead As String
    Dim currentwrite As String
    
    Do Until readtext.AtEndOfStream = True
        CurrentRead = readtext.Read(1)
    Loop
End Sub
