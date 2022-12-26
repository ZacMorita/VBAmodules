Attribute VB_Name = "Module1"
Option Explicit

'This Class is written to work with both "Early Binding" and "Late Binding" _
    of the "Microsoft Word 16.0 Object Library". _
    If compilation fails with message: _
    "Compile Error: User-Defined Type Not Defined" then, in the _
    Visual Basic Editor Go to Tools >> References... >> Then check the box _
    next to the mentioned library. Or, if Late Binding is required, _
    manually change "#Const EarlyBinding = True" to "#Const EarlyBinding = False" _
    Early Binding is exponentially faster and less error prone than Late Binding.

#Const EarlyBinding = True '<--- This is manually adjufted for Early and Late Binding

Public Sub ManipulateWord()
    'IMPORTANT NOTE: _
        If compilation fails with message "Compile Error: User-Defined Type Not Defined" _
        Please see "EarlyBinding" Comment at the top of this Class Module.
    'EarlyBinding is a pre-compile constant at the top of this module
    #If EarlyBinding Then 'This is a pre-compile conditional
        Dim WordApp As Word.Application
        Dim doc As Word.Document
        Dim StoryRange As Word.Range
    #Else
        Dim WordApp As Object
        Dim doc As Object
        Dim StoryRange As Object
    #End If
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim value1 As Double
    Dim value2 As Double
    Dim mathing As Double

    'Error handling at bottom of sub, Ensures word.app closes out proper on error
    On Error GoTo AbortSub
    Set wb = ThisWorkbook
    Set ws = wb.Sheets(1)

    'Fill number variables
    value1 = ws.Range("A1").Value
    value2 = ws.Range("B1").Value
    If Not value1 = 0 Then
        mathing = (value2 / value1) - 1
    Else
        MsgBox "Division by zero: check value1"
        GoTo AbortSub
    End If

    'EarlyBinding is a pre-compile constant at the top of this module
    #If EarlyBinding Then 'This is a pre-compile conditional
        Set WordApp = New Word.Application
    #Else
        Set WordApp = CreateObject("Word.Application")
    #End If
    
    With WordApp
        .Visible = True
        'setting Doc this way ensures you can use it later in the sub
        Set doc = .Documents.Add
    End With
    
    'doc.Content represents the Range that makes up the content of Document
    Set StoryRange = doc.Content
    
    With StoryRange
        'set Document.Content.Font properties
        With .Font
            .Size = 9
            .Name = "Arial"
        End With
        'Write a line with variables into the .Contents of the new Document
        .InsertAfter "Revenues changed by " & _
                     Format(mathing, "0.00%") & _
                     " from " & _
                     value2 & _
                     " to " & _
                     value1
        .InsertAfter vbCrLf 'vbCrLf is a constant that represents a NewLine character
        .InsertAfter "COGS went from" 'Write more things
        'Another example
        .InsertAfter vbCrLf & vbCrLf & "Test" & vbCrLf & "More Test" & vbCrLf & vbCrLf & "Even More"
    End With
    Exit Sub 'end sub without running into the error-handlers script
AbortSub: 'error handling
    If Not doc Is Nothing Then
        doc.Close False ' close without saving
    End If
    If Not WordApp Is Nothing Then
        'If Word.Application not ".Visible = True" and not ".quit" it _
            will be stuck loaded in the background even after closing excel
        WordApp.Quit False 'Close without saving
    End If
    Set doc = Nothing 'Not always necessary house-cleaning
    Set WordApp = Nothing 'Not always necessary house-cleaning
    If Err.Number <> 0 Then
        'Raise the error after closes so the user knows something went wrong
        Err.Raise Err.Number
    End If
End Sub

