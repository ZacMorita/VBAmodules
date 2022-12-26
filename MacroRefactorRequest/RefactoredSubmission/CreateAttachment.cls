VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BasePath As String = "C:\Examples\Example\"
Private Const AttachmentExtension As String = ".xlsx"
Private Const FileExistsMessage As String = vbCrLf & "already exists, would you like to overwrite it?" & _
                                            vbCrLf & vbCrLf & "Press Yes to overwrite with a new report and attach it to the email." & _
                                            vbCrLf & vbCrLf & "Press No to attach existing report to the email."

Private NewFileName As String
Private m_AttachmentPath As String

Private Sub Class_Initialize()
    'VBA doesn't support class cunstructors with arguments and because _
        m_AttchmentPath is critical to operation and value-sensitive. It's _
        value is "Hard Coded" and set upon class initialization.
    NewFileName = "Daily Passdown " & Format(Date, "mm-dd-yyyy")
    m_AttachmentPath = BasePath & NewFileName & AttachmentExtension
End Sub

Public Property Get AttachmentPath() As String
    AttachmentPath = m_AttachmentPath
End Property

Public Sub GeneratePassdownAttachment()
    Dim ThisWB As Workbook
    Dim NewWB As Workbook
    Dim ThatWS As Worksheet
    Dim PassDownRange As Range
    Dim FileExisted As Boolean
    Dim MessageResponse As VbMsgBoxResult
    Dim testing As Integer

    FileExisted = CustomFunctions.FileExists(m_AttachmentPath)
    If FileExisted Then
        'FileExistsMessage is Constant at the top of this module.
        MessageResponse = MsgBox(m_AttachmentPath & FileExistsMessage, vbYesNo, "Attachment Aleardy Created")
        If MessageResponse = vbYes Then
            'Destroy existing file to be replaced below.
            Kill m_AttachmentPath
        End If
    End If

    If FileExisted = False Or MessageResponse = vbYes Then
        'Prevent screenflash and improve performance
        Application.ScreenUpdating = False

        Set ThisWB = ThisWorkbook
        Set NewWB = Application.Workbooks.Add
        Set ThatWS = NewWB.Sheets(1)
        'The following range can be found in the Excel Application Name Manager by pressing CTRL + F3
        Set PassDownRange = ThisWB.Sheets("Daily Passdown").Range("DailyPassDownRange")

        'Copy the Passdown working area
        PassDownRange.Copy

        'Paste it's contents and format in the NewWB
        With ThatWS.Range("A1")
            .PasteSpecial xlPasteColumnWidths
            .PasteSpecial xlPasteValues
            .PasteSpecial xlPasteFormats
            .Select
        End With
        
        With ThatWS.PageSetup
            .Orientation = xlLandscape
            .PrintArea = ThatWS.UsedRange.Address
            .Zoom = False
            .FitToPagesTall = 1
            .FitToPagesWide = 1
        End With

        'Ends the Copy/Cut/Paste mode
        Application.CutCopyMode = False

        'Saves and closes NewWB.
        With NewWB
            .SaveAs Filename:=m_AttachmentPath, _
                    FileFormat:=xlOpenXMLWorkbook, _
                    ConflictResolution:=xlLocalSessionChanges, _
                    CreateBackup:=False
            .Close False
        End With

        'reset ScreenUpdating property
        Application.ScreenUpdating = True
    End If
End Sub

