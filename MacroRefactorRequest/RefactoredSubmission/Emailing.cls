VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Emailing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This Class is written to work with both "Early Binding" and "Late Binding" _
    of the "Microsoft Outlook 16.0 Object Library". _
    If compilation fails with message "Compile Error: User-Defined Type Not Defined" _
    Go to Tools >> References... >> Then check the box next to the mentioned library. _
    Or, if Late Binding is required, manually change "#Const EarlyBinding = True" to "#Const EarlyBinding = False" _
    Early Binding is exponentially faster and less error prone than Late Binding.

#Const EarlyBinding = True '<--- This is a manually adjufted setting for Early and Late Binding

Private Const OlNotOpenMessage As String = "Outlook not open." & vbCrLf & vbCrLf & _
                                           "Email preview window unavailable." & vbCrLf & vbCrLf & _
                                           "Press Okay to send Daily Passdown email with no preview." & vbCrLf & vbCrLf & _
                                           "Press Cancel to abort email. (No draft will be made)" & vbCrLf & vbCrLf & _
                                           "If sent, email will be sent to: "

#If EarlyBinding Then
    Private OLApp As Outlook.Application
#Else
    Private OLApp As Object
#End If

Private UsingExistingOLApp As Boolean

Private Sub Class_Initialize()
    'IMPORTANT NOTE: _
        If compilation fails with message "Compile Error: User-Defined Type Not Defined" _
        Please see "EarlyBinding" Comment at the top of this Class Module.
    
    On Error GoTo ErrHandler
    'Will attempt to use an open Outlook application to generate emails. _
        If no Outlook is open, an error 492 will trigger (see "ErrHandler:" at end of this sub).
    Set OLApp = GetObject(, "Outlook.Application") '[Error 429 if not open]

    If OLApp Is Nothing Then
        'If no Outlook was open for the above "GetObject" a new outlook will be ran from the background.
        #If EarlyBinding Then
            Set OLApp = New Outlook.Application
        #Else
            Set OLApp = CreateObject("Outlook.Application")
        #End If
        UsingExistingOLApp = False 'Boolean for SendEmail() to handle .send/.display
    Else
        UsingExistingOLApp = True 'Boolean for SendEmail() to handle .send/.display
    End If
    
    Exit Sub
ErrHandler:
    If Err.Number = 429 Then
        Err.Clear 'This clears the error because it was expected/intended. (if outlook is not open)
        Resume Next 'This returns program flow to the line after the error occurred to create a new Outlook object.
    ElseIf Err.Number <> 0 Then
        'if error for reason other than 429 (outlook not open) then show the error and end the sub.
        Err.Raise Err.Number
        Exit Sub
    End If
    
End Sub

Private Sub Class_Terminate()
    Set OLApp = Nothing
End Sub

Public Sub SendEmail(ByRef Reciever As String, _
                     ByRef CCd As String, _
                     ByRef SubjectLine As String, _
                     ByRef EmailBody As String, _
                     ByRef Attachment As String)
                
    'IMPORTANT NOTE: _
        If compilation fails with message "Compile Error: User-Defined Type Not Defined" _
        Please see "EarlyBinding" Comment at the top of this Class Module.
    #If EarlyBinding Then
        Dim Message As Outlook.MailItem
    #Else
        Dim Message As Object
    #End If

    Set Message = OLApp.CreateItem(0) 'olMailItem = 0 (Enum of Outlook.OlItemType)

    'Construct Message from arguments
    With Message
        .BodyFormat = 2 'olFormatHTML = 2 (Enum of Outlook.OlBodyFormat)
        .To = Reciever
        .CC = CCd
        .Subject = SubjectLine
        .HTMLBody = EmailBody
        .Attachments.Add Attachment
        If Not UsingExistingOLApp Then
            'OlNotOpenMessage is Constant at the top of this module.
            If MsgBox(OlNotOpenMessage & Reciever & " " & CCd, vbOKCancel, "Outlook not open") = vbOK Then
                .send 'Send without preview - will send if outlook is closed
            End If
        ElseIf UsingExistingOLApp Then
            .Display 'Show preview - this method wont send if outlook is closed
        End If
    End With
    
    Set Message = Nothing
End Sub
