Attribute VB_Name = "Module1"
Option Explicit

Sub PrintPageSetUpWithPerformanceEnhancers()
    Dim RememberPrintCommunication As Boolean
    Dim RememberDisplayAlerts As Boolean

    'NOTE: .PrintCommunication can be problematic when set to True _
        as all page set up changes will be sent to the active printer _
        at that time. If Runtime Error 1004 - Object Method Failed _
        then communication to the active printer failed. Then which _
        PrintPageSetUpAlternative() [below] should be used.

    'Hold Booleans to set back at end of Sub
    RememberPrintCommunication = Application.PrintCommunication
    RememberDisplayAlerts = Application.DisplayAlerts
    'Disable Printer communication for faster page setup
    Application.PrintCommunication = False
    'Disable displayed alerts to prevent setup alerts due to _
        printer communication being disabled.
    Application.DisplayAlerts = False
    
    'page setup
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .PrintArea = ActiveSheet.UsedRange.Address
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = False
        .Zoom = False 'FitToPageTall & Wideare ignored is Zoom is True
        .FitToPagesTall = False
        .FitToPagesWide = 1
    End With
    
    'restore changed settings
    Application.PrintCommunication = Application.PrintCommunication
    Application.DisplayAlerts = RememberDisplayAlerts
End Sub
'-------------------------------------------------------------
Sub PrintPageSetUpAlternative()
        'page setup
        With ActiveSheet.PageSetup
            .Orientation = xlLandscape
            .PrintArea = ActiveSheet.UsedRange.Address
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = False
            .Zoom = False 'FitToPageTall & Wideare ignored is Zoom is True
            .FitToPagesTall = False
            .FitToPagesWide = 1
        End With
End Sub



