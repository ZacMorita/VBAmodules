Attribute VB_Name = "Module1"

Sub DailyPassdown()

Application.ScreenUpdating = False

    Sheets("Board").Select

'281
    Range("F13:H15").Select
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'282
    Range("J13:L15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'283
    Range("B25:D27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'285
    Range("F25:H27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'286
    Range("J25:L27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    
'Cleaning up
    Sheets("Board").Select
    Application.CutCopyMode = False 'test
    Range("F13").Select
    Sheets("Daily Passdown").Select
        
'Passdown Sheet Hidden table
    Range("V3:X17").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("W2:W17").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    
'Pasting information
    Range("B19:C33").Select
    Selection.ClearContents     'Clear
    Range("V3:V17").Select
    Selection.Copy
    Range("B19:C19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Range("F19:F33").Select
    Selection.ClearContents     'Clear
    Range("W3:W17").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("F19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Range("G19:G33").Select
    Selection.ClearContents     'Clear
    Range("X3:X17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
'Cleaning up
    Range("V3:X17").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    
'Staffing Hours
    Sheets("Input Data").Select
    Range("L12:M16").Select
    Selection.Copy
    Range("N17").Select
    Sheets("Daily Passdown").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'Cleaning up Pt. 2
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Range("B19:C19").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub Update()

Application.ScreenUpdating = False

Sheets("Capacity Query").Select
Range("A1").Select
Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
Range("DataMart_viewName_TemporaryCapacityReservations[[#Headers],[Value.resource]]").Select

Sheets("Pivot Table - Loaded Hours").Select
Range("A1").Select
ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh

Sheets("Input Data").Select
Range("N3").Value = "made you look"
Range("N3").Value = ""

Application.ScreenUpdating = True

End Sub

Sub NewWorkBook()

Dim currentdate As String
Dim newfilename As String

Sheets("Input Data").Select
Range("T1").Select
Selection.Copy
Range("N1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
currentdate = Range("N1").Value
newfilename = "Daily Passdown " & currentdate
Workbooks.Add.SaveAs Filename:=newfilename

End Sub

Sub Transfer()

    Sheets("Daily Passdown").Select
    Range("A1:H33").Select
    Selection.Copy
    Windows("Daily Passdown 04-20-2022.xlsx").Activate      'Is manually entered for this example
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    Range("A1").Select
    Windows("Andon Board V10.xlsm").Activate
    Application.CutCopyMode = False
    Range("A1").Select

End Sub

Sub ExportFinal()

'''Email var
Dim emailApplication As Object
Dim emailItem As Object

'''Pivot table var
Dim tablerange As String
Dim initial As String

'''Workbook var
Dim currentdate As String
Dim newfilename As String
Dim send As String

'Set Email var
Set emailApplication = CreateObject("Outlook.Application")
Set emailItem = emailApplication.CreateItem(0)

'''Update routine
Application.ScreenUpdating = False

'Refresh Query
Sheets("Capacity Query").Select
Range("A1").Select
Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
Range("DataMart_viewName_TemporaryCapacityReservations[[#Headers],[Value.resource]]").Select

'Refresh pivot table
Sheets("Pivot Table - Loaded Hours").Select
Range("A1").Select
ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh

'Section is not entirely necessary, was used to update clock in an individual sub
Sheets("Input Data").Select
Range("N3").Value = "made you look"
Range("N3").Value = ""

'''DailyPassdown routine
    Sheets("Board").Select

'281
    Range("F13:H15").Select
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'282
    Range("J13:L15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'283
    Range("B25:D27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'285
    Range("F25:H27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'286
    Range("J25:L27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    
'Cleaning up
    Sheets("Board").Select
    Application.CutCopyMode = False
    Range("F13").Select
    Sheets("Daily Passdown").Select
        
'Passdown Sheet Hidden table
    Range("V3:X17").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("W2:W17").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    
'Pasting information
    Range("B19:C33").Select
    Selection.ClearContents     'Clear
    Range("V3:V17").Select
    Selection.Copy
    Range("B19:C19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Range("F19:F33").Select
    Selection.ClearContents     'Clear
    Range("W3:W17").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("F19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Range("G19:G33").Select
    Selection.ClearContents     'Clear
    Range("X3:X17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
'Cleaning up
    Range("V3:X17").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    
'Staffing Hours
    Sheets("Input Data").Select
    Range("L12:M16").Select
    Selection.Copy
    Range("N17").Select
    Sheets("Daily Passdown").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'Cleaning up Pt. 2
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Range("B19:C19").Select

Application.ScreenUpdating = True

'''Email selection range
Sheets("Pivot Table - Loaded Hours").Select
Cells(2, 1).Select
Selection.End(xlToRight).Select
Selection.End(xlDown).Select
initial = ActiveCell.Address
tablerange = "A1" & ":" & initial
Sheets("Input Data").Select
Range("N17").Select

'''Email content
emailItem.To = "jane.doe@example.com"
emailItem.CC = "john.doe@example.com"     'Range("enter email here").Value
emailItem.BCC = ""    'Range("enter email here").Value
emailItem.Subject = "Hours from " & Range("N2").Value     'N2 calls date/time with =now() function
emailItem.HTMLBody = "Hello!" & "<br>" & "<br>" & "The attached file/table was last updated " & Range("N2").Value & "<br>" & "<br>" & ConvertRangeToHTMLTable(Sheet6.Range(tablerange)) & "<br>" & "<br>"
'  "<br>" is HTML line break, vbLf for non HTML body (do not put vbLf in quotes)

'''Email attachment

'Create new workbook                                                         EXPERIMENTAL WITH DAILY PASSDOWN S1
Sheets("Input Data").Select
Range("S1").Select
Selection.Copy
Sheets("Daily Passdown").Select
Range("W1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
currentdate = Range("W1").Value                     'currentdate location
newfilename = "Daily Passdown " & currentdate       'newfilename location
Workbooks.Add
ChDir "C:\EXAMPLE\Daily Passdown & Andon Board"
ActiveWorkbook.SaveAs Filename:="C:\EXAMPLE\Daily Passdown & Andon Board\" & newfilename & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False 'S:\ format is shared drive. a C":\ would be local
Windows("Andon Board V10.6.xlsm").Activate             'Calls Andon file name manually  !OFI!  make it a variable to update versions easily
Workbooks.Open Filename:="C:\EXAMPLE\Daily Passdown & Andon Board\Andon Board V10.6.xlsm"      'Calls Andon file name manually

'Add passdown info copy to new workbook
Sheets("Daily Passdown").Select
Range("A1:H33").Select
Selection.Copy
Workbooks.Open "C:\EXAMPLE\Daily Passdown & Andon Board\" & newfilename & ".xlsm"
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
ActiveSheet.Paste
Range("A1").Select
Sheets("Sheet1").Name = newfilename
ActiveWorkbook.Save

'Back to Andon Board
Windows("Andon Board V10.6.xlsm").Activate            'Calls Andon file name manually
Application.CutCopyMode = False
Range("A1").Select

'Calls workbook from files
send = "C:\EXAMPLE\Daily Passdown & Andon Board\" & newfilename & ".xlsm"
emailItem.Attachments.Add send   'send is variable name line above

'''Shows email before sending, use .Send to automatically send
emailItem.Display

                   'Primary efficiency spot

'Cleaning up
Sheets("Daily Passdown").Select
Range("A1").Select

'''Clear
Set emailApplication = Nothing
Set emailItem = Nothing

End Sub

Sub primaryEfficiency()

'''Update Primary Efficiency table
Sheets("Primary Efficiency").Select
Range("S2:W7").Select
Selection.Copy
Range("S11:W16").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("S11:W16").Select
Selection.Copy
Range("A2").Select
Selection.End(xlDown).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("Daily Passdown").Select

''' 285 Location addon

Sheets("Input Data").Select

Dim arrayNames(1 To 4) As Variant

For i = 1 To 4

    arrayNames(i) = 0
    
Next i


For i = 1 To 4
    
    Sheets("Input Data").Select
    arrayNames(i) = Range("S" & i + 11).Value * Range("M3").Value
    Sheets("285 Efficiency Breakdown").Select
    Range("U" & i + 1).Value = arrayNames(i)
    
Next i

Range("S2:U6").Select
Selection.Copy
Range("A2").Select
Selection.End(xlDown).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Application.CutCopyMode = False
Selection.NumberFormat = "m/d/yyyy"
Range("A2").Select
    
Sheets("Daily Passdown").Select

MsgBox ("Success!")

End Sub

Sub ExportTesting()

'''Email var
Dim emailApplication As Object
Dim emailItem As Object

'''Pivot table var
Dim tablerange As String
Dim initial As String

'''Workbook var
Dim currentdate As String
Dim newfilename As String
Dim send As String

Set emailApplication = CreateObject("Outlook.Application")
Set emailItem = emailApplication.CreateItem(0)

'''Update routine
Application.ScreenUpdating = False

Sheets("Capacity Query").Select
Range("A1").Select
Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
Range("DataMart_viewName_TemporaryCapacityReservations[[#Headers],[Value.resource]]").Select

Sheets("Pivot Table - Loaded Hours").Select
Range("A1").Select
ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh

Sheets("Input Data").Select
Range("N3").Value = "made you look"
Range("N3").Value = ""

'''DailyPassdown routine
    Sheets("Board").Select

'281
    Range("F13:H15").Select
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'282
    Range("J13:L15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'283
    Range("B25:D27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'285
    Range("F25:H27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Board").Select
    
'286
    Range("J25:L27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Daily Passdown").Select
    Range("V15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    
'Cleaning up
    Sheets("Board").Select
    Application.CutCopyMode = False
    Range("F13").Select
    Sheets("Daily Passdown").Select
        
'Passdown Sheet Hidden table
    Range("V3:X17").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("W2:W17").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    
'Pasting information
    Range("B19:C33").Select
    Selection.ClearContents     'Clear
    Range("V3:V17").Select
    Selection.Copy
    Range("B19:C19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Range("F19:F33").Select
    Selection.ClearContents     'Clear
    Range("W3:W17").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("F19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Range("G19:G33").Select
    Selection.ClearContents     'Clear
    Range("X3:X17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
'Cleaning up
    Range("V3:X17").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    
'Staffing Hours
    Sheets("Input Data").Select
    Range("L12:M16").Select
    Selection.Copy
    Range("N17").Select
    Sheets("Daily Passdown").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'Cleaning up Pt. 2
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Range("B19:C19").Select
    
'''Update Primary Efficiency table
Sheets("Primary Efficiency").Select
Range("S2:W7").Select
Selection.Copy
Range("A2").Select
Selection.End(xlDown).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Application.ScreenUpdating = True

'''Email selection range
Sheets("Pivot Table - Loaded Hours").Select
Cells(2, 1).Select
Selection.End(xlToRight).Select
Selection.End(xlDown).Select
initial = ActiveCell.Address
tablerange = "A1" & ":" & initial
Sheets("Input Data").Select
Range("N17").Select

'''Email content
emailItem.To = "jane.doe@example.com"
emailItem.CC = "john.doe@example.com"     'Range("enter email here").Value
emailItem.BCC = ""    'Range("enter email here").Value
emailItem.Subject = "Hours from " & Range("N2").Value
emailItem.HTMLBody = "Hello!" & "<br>" & "<br>" & "The attached file/table was last updated " & Range("N2").Value & "<br>" & "<br>" & ConvertRangeToHTMLTable(Sheet3.Range(tablerange)) & "<br>" & "<br>"

'''Email attachment

'Create new workbook
Sheets("Input Data").Select
Range("T1").Select
Selection.Copy
Range("N1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
currentdate = Range("N1").Value
newfilename = "Daily Passdown " & currentdate
Workbooks.Add
ChDir "C:\EXAMPLE\Daily Passdown & Andon Board"
ActiveWorkbook.SaveAs Filename:="C:\EXAMPLE\Daily Passdown & Andon Board\" & newfilename & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
Windows("Andon Board V10.6.xlsm").Activate
Workbooks.Open Filename:="C:\EXAMPLE\Daily Passdown & Andon Board\Andon Board V10.6.xlsm"

'Add passdown info copy to new workbook
Sheets("Daily Passdown").Select
Range("A1:H33").Select
Selection.Copy
Workbooks.Open "C:\EXAMPLE\Daily Passdown & Andon Board\" & newfilename & ".xlsm"
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
ActiveSheet.Paste
Range("A1").Select
Sheets("Sheet1").Name = newfilename
ActiveWorkbook.Save

'Back to Andon Board
Windows("Andon Board V10.xlsm").Activate
Application.CutCopyMode = False
Range("A1").Select

'Calls workbook from files
send = "C:\EXAMPLE\Daily Passdown & Andon Board\" & newfilename & ".xlsm"
emailItem.Attachments.Add send

'''Shows email before sending, use .Send to automatically send
emailItem.Display

'''Clear
Set emailApplication = Nothing
Set emailItem = Nothing

End Sub

Public Function ConvertRangeToHTMLTable(rInput As Range) As String

    Dim rRow As Range
    Dim rCell As Range
    Dim strReturn As String
    
    'Define table format and font
    strReturn = "<Table border='1' cellspacing='0' cellpadding='7' style='border-collapse:collapse;border:none'>  "
    
    'Loop through each row in the range
    For Each rRow In rInput.Rows
    
        'Start new html row
        strReturn = strReturn & " <tr align='Center'; style='height:10.00pt'> "
        
        For Each rCell In rRow.Cells
        
            'If it is row 1 then it is header row that need to be bold
            If rCell.Row = 1 Then
            
                strReturn = strReturn & "<td valign='Center' style='border:solid windowtext 1.0pt; padding:0cm 5.4pt 0cm 5.4pt;height:1.05pt'><b>" & rCell.Text & "</b></td>"
                
            Else
            
                strReturn = strReturn & "<td valign='Center' style='border:solid windowtext 1.0pt; padding:0cm 5.4pt 0cm 5.4pt;height:1.05pt'>" & rCell.Text & "</td>"
                
            End If
            
        Next rCell
        
        'End a row
        strReturn = strReturn & "</tr>"
        
    Next rRow
    
    'Close the font tag
    strReturn = strReturn & "</font></table>"
    
    'Return html format
    ConvertRangeToHTMLTable = strReturn
    
End Function


