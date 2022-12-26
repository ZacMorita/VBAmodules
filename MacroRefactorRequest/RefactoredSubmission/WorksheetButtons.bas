Attribute VB_Name = "WorksheetButtons"
Option Explicit

'This module contains subs that are tied to the buttons on (the) worksheet(s) _
    within ThisWorkbook.

'NOTE: Email strings declared as constants for legacy reasons.
'NOTE: Email strings shouldn't be hard coded constants. Consider decalring e_To _
    and e_CC as local variables in GenerateEmailFromWorkbookData() _
    and maybe storing them in a Worksheet or Named Range else maybe a user _
    input prompt for ease of accesibility and user awareness.
Public Const e_To As String = "jane.doe@example.com"
Public Const e_CC As String = "john.doe@example.com"
Public Const e_Subject As String = "Hours from "
Public Const e_Body As String = "Hello!<br><br>The attached file/table was last updated "
Public Const e_Body_br As String = "<br><br>"

Public Sub SetConnectionPropertiesAndRefreshAll()
    'This sub executes the function of the "Refresh All" button on _
        the "Data" ribbon from the Excel App after ensuring connection _
        Properties are set. The connection properties currently set to _
        maximize Workbook load time efficiency and prioritizes refreshing _
        only for the sake of other automation. If regularly updated data is _
        necessary, these settings should be changed.
    Dim wb As Workbook
    Dim cons As Connections
    Dim con As WorkbookConnection

    Set wb = ThisWorkbook
    Set cons = wb.Connections

    For Each con In cons
        If con.Type = xlConnectionTypeOLEDB Then
            With con.OLEDBConnection
                'Enable and disable all refreshing of current connection, ignores all other settings
                .EnableRefresh = True

                'Disables asynchronous refreshing, this should ensure the refresh completes before _
                    emailing reports occurs (when ran in conjuntion with other automation processes)
                .BackgroundQuery = False

                '0 disables RefreshPeriod, positive numbers Enable it and set the period in Minutes
                .RefreshPeriod = 30

                'Workbook opens faster with False, set to true if fresh data on open is necessary
                .RefreshOnFileOpen = True
            End With

            'Ensures the connections work with "Refresh All". This should be changed to incorporate _
                conditionals if selective refreshing for ignoring certain connections is required. _
                simple example is included and commented out.
            'If Not con.Name = "A query I wish to skip" Then
                con.RefreshWithRefreshAll = True
            'Else
                'con.RefreshWithRefreshAll = False
            'End If

        End If
    Next con

    wb.RefreshAll
    
    For Each con In wb.Connections
        If con.Type = xlConnectionTypeOLEDB Then
            With con.OLEDBConnection
                .BackgroundQuery = True
            End With
        End If
    Next con

End Sub

Public Sub TransferDailyPassdown()
    'This sub calls subs from within this module to move data from the "Board" sheet _
        and the "Input Data" sheet to the "Daily Passdown" sheet.
    
    'a sub from this module
    WorksheetButtons.TransferDailyPassdownPart1
    'a sub from this module
    WorksheetButtons.TransferDailyPassdownPart2
End Sub

Public Sub TransferDailyPassdownPart1()
    'NOTE: This sub is called in conjunction with the sub "TransferDailyPassdownPart2()" _
        from the sub "TransferDailyPassdown()"
    'This sub moves all the user input values from Sheets("Board") _
        to the "Line Down Tracking" section of Sheets("Daily Passdown")
    Dim BoardAreas As Areas
    Dim LDTRange As Range
    Dim rng As Range
    Dim rw As Range
    Dim i As Integer
    
    i = 1
    'The following ranges can be found in the Excel Application Name Manager by pressing CTRL + F3
    Set BoardAreas = ThisWorkbook.Sheets("Board").Range("BoardAreas").Areas
    Set LDTRange = ThisWorkbook.Sheets("Daily Passdown").Range("LDTRange")
    
    LDTRange.ClearContents
    
    For Each rng In BoardAreas
        For Each rw In rng.Rows
            'If "Item"/"Item Number" ISBLANK then skip row.
            If Not IsEmpty(rw.Cells(1, 1)) Then
                With LDTRange
                    'column values at offset interval due to range addressing of merged cells.
                    .Cells(i, 1) = rw.Cells(1, 1)
                    .Cells(i, 5) = rw.Cells(1, 2)
                    .Cells(i, 6) = rw.Cells(1, 3)
                End With
                i = i + 1
            End If
        Next rw
    Next rng
End Sub

Public Sub TransferDailyPassdownPart2()
    'NOTE: This sub is called in conjunction with the sub "TransferDailyPassdownPart1()" _
        from the sub "TransferDailyPassdown()"
    'This sub copies the values in the 2nd and 3rd cloumn of BackflushInputRange on the _
        "Input Data" sheet to the DailyPassdownHours range on the "Daily Passdown" Sheet.
    Dim BackflushInputRange As Range
    Dim DailyPassdownHours As Range
    Dim ArrRange As Variant

    'The following ranges can be found in the Excel Application Name Manager by pressing CTRL + F3
    Set BackflushInputRange = ThisWorkbook.Sheets("Input Data").Range("BackflushInputRange")
    Set DailyPassdownHours = ThisWorkbook.Sheets("Daily Passdown").Range("DailyPassdownHours")
    'Assign backflushhours values to an array
    ArrRange = Range(BackflushInputRange.Cells(1, 2), BackflushInputRange.Cells(5, 3)).Value
    'Put the array of values in thedailypassdown range.
    DailyPassdownHours.Value = ArrRange
End Sub

Public Sub GenerateEmailFromWorkbookData()
    'This sub compiles the contents of an email from Constants at the top _
        of this module and data in ThisWorkbook. Then uses the 'CreateAttachment' _
        class from this project to create a new report Workbook to be used as _
        an email attachment. Then uses the 'Emailing' class from this project to _
        create a Microsoft Outlook email, populates it, then sends it with a user prompt.
    Dim wb As Workbook
    Dim email As Emailing
    Dim CreateAttach As CreateAttachment
    Dim PivTab As PivotTable
    Dim PivTabRng As Range
    Dim EmailBodyTable As Range

    Set wb = ThisWorkbook
    'Emailing is a Class from this project
    Set email = New Emailing
    'CreateAttachment is a Class from this project
    Set CreateAttach = New CreateAttachment

    Set PivTab = wb.Sheets("Pivot Table - Loaded Hours").PivotTables("PivotTable1")

    'Set EmailBodyTable to adusted size and range using the dimensions of PivTab.TableRange1
    Set EmailBodyTable = PivTab.TableRange1.Offset(1).Resize(PivTab.TableRange1.Rows.Count - 1)
    
    'MoreThanAnHourAgo is a custom function from this project _
        This condition is commented out for a quick available option.
    'If CustomFunctions.MoreThanAnHourAgo(PivTab.RefreshDate) Then
        'a sub from this module
        WorksheetButtons.SetConnectionPropertiesAndRefreshAll
    'End If

    'a sub from this module
    WorksheetButtons.TransferDailyPassdown

    'Create the passdown attachment using path stored in the ReadOnly-Property _
        CreateAttachment.AttachmentPath which is filled when CreatAttachment _
        Initializes. Changes can be made from within the CreateAttachment Class Module.
    CreateAttach.GeneratePassdownAttachment
    
    'Fill custom emailing method with Constants (found at top of this module), emailbody _
         using custom function, and attachment from custom class
    email.SendEmail e_To, _
                    e_CC, _
                    e_Subject & Now, _
                    e_Body & _
                        Now & _
                        e_Body_br & _
                        CustomFunctions.ConvertRangeToHTMLTable(EmailBodyTable) & _
                        e_Body_br, _
                    CreateAttach.AttachmentPath

    Set email = Nothing
End Sub

Public Sub TransferPrimaryEfficiency()
    'This sub moves data from two areas in the "Input Data" sheet. To two other _
        areas in Thisworkbook while applying percentile math to the later.

    'a sub from this module
    WorksheetButtons.PrimaryEfficiencyPart1
    'a sub from this module
    WorksheetButtons.PrimaryEfficiencyPart2

    'Indicate to the user that actions were performed. _
        Not necessarly indicative of successful transfer.
    MsgBox "Reached end of sub.", vbOKOnly, "Sub Complete"

End Sub

Private Sub PrimaryEfficiencyPart1()
    'NOTE: This sub is called in conjunction with the sub "PrimaryEfficiencyPart2()" _
        from the sub "TransferPrimaryEfficiency()"
    'This sub moves data from "Backflush hours and people assigned" area to the table _
        "TableEfficiency2" on the "Primary Efficiency" sheet.
    Dim wb As Workbook
    Dim BackflushInputRange As Range
    Dim CurrentRow As Range
    Dim LookRange As Range
    Dim FoundRange As Range
    Dim PrimeEffTable As ListObject
    Dim i As Integer
    
    Set wb = ThisWorkbook
    'The following range can be found in the Excel Application Name Manager by pressing CTRL + F3
    Set BackflushInputRange = wb.Sheets("Input Data").Range("BackflushInputRange")
    Set PrimeEffTable = wb.Sheets("Primary Efficiency").ListObjects("TableEfficiency2")
    Set LookRange = PrimeEffTable.ListColumns(1).DataBodyRange
    Set FoundRange = LookRange.Find(What:=Date, LookIn:=xlValues)

    Do Until FoundRange Is Nothing
        FoundRange.EntireRow.Delete
        Set FoundRange = LookRange.FindNext
    Loop
    For Each CurrentRow In BackflushInputRange.Rows
        With PrimeEffTable.ListRows.Add
            'Dates the line
            .Range(1, 1) = Date
            'Names the "Area" using the number found at the begining of "### Hours" _
                part of range "BackflushInputRange" (Example: "281 Hours" -> "DS 281")
            .Range(1, 2) = "DS " & Left(CurrentRow.Cells(1, 1).Value, 3)
            'Transfers row values to table
            For i = 3 To 5
                .Range(1, i) = CurrentRow.Cells(1, i - 1).Value
            Next i
        End With
    Next CurrentRow
End Sub

Private Sub PrimaryEfficiencyPart2()
    'NOTE: This sub is called in conjunction with the sub "PrimaryEfficiencyPart1()" _
        from the sub "TransferPrimaryEfficiency()"
    'This sub moves data from "285 Location and # People" to the bottom of the values _
        on the sheet "285 Efficiency Breakdown" while multiplying "# People" _
        values by the percentage stored in the first cell of "Efficiency" _
        on the "Input Data" sheet.
    Dim wb As Workbook
    Dim ws285BD As Worksheet
    Dim Loc285People As Range
    Dim CurrentRow As Range
    Dim LookRange As Range
    Dim FoundRange As Range
    Dim EfficiencyPercent As Double
    Dim NthRowws285BD As Integer
    Dim i As Integer

    Set wb = ThisWorkbook
    'The following ranges can be found in the Excel Application Name Manager by pressing CTRL + F3
    With wb.Sheets("Input Data")
        Set Loc285People = .Range("Loc285People")
        EfficiencyPercent = .Range("EfficiencyPercent").Value
    End With
    Set ws285BD = wb.Sheets("285 Efficiency Breakdown")
    NthRowws285BD = ws285BD.Cells(ws285BD.Rows.Count, 1).End(xlUp).Row + 1
    
    Set LookRange = ws285BD.Range(ws285BD.Cells(1, 1), ws285BD.Cells(NthRowws285BD, 1))
    Set FoundRange = LookRange.Find(What:=Date, LookIn:=xlValues)

    Do Until FoundRange Is Nothing
        FoundRange.EntireRow.Delete
        Set FoundRange = LookRange.FindNext
        NthRowws285BD = NthRowws285BD - 1
    Loop

    For Each CurrentRow In Loc285People.Rows
        'Dates the line
        ws285BD.Cells(NthRowws285BD, 1).Value = Date
        'Labels Line
        ws285BD.Cells(NthRowws285BD, 2).Value = CurrentRow.Cells(1, 1).Value
        'Percentile math using range "EfficiencyPercent"
        ws285BD.Cells(NthRowws285BD, 3).Value = CurrentRow.Cells(1, 2).Value * EfficiencyPercent
        'increment to the next lowest row.
        NthRowws285BD = NthRowws285BD + 1
    Next CurrentRow
End Sub
