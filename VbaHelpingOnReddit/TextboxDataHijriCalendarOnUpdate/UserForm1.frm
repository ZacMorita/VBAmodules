VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    Dim RememberedCalendar As VbCalendar
    Dim InputDateStr As String
    Dim OutputDate As Date
    Dim Millennia As Long

    RememberedCalendar = VBA.Calendar
    InputDateStr = Me.TxtDate.Value
    'Requester asked short dates like "5/22" default Hijri
    If Not Len(InputDateStr) <= 5 Then
        Millennia = Year(CDate(InputDateStr))
    End If

    If Millennia >= 1900 Then
        VBA.Calendar = vbCalGreg
    Else
        VBA.Calendar = vbCalHijri
    End If

    OutputDate = CDate(InputDateStr)

    If VBA.Calendar = vbCalGreg Then
        VBA.Calendar = vbCalHijri
    End If
    Me.TxtDate.Value = Format(OutputDate, "yyyy-mm-dd")
    VBA.Calendar = RememberedCalendar
End Sub
