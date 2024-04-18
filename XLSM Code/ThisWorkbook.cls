VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' --------------------------------------------
' Project: Attack on Europe 2022
' Made by: Matthew Borcz
' GH Page: https://github.com/MBrcz
' --------------------------------------------

' -------------------------------------------------
' ------------------ UTILS ------------------------
' -------------------------------------------------

Public Sub ExecuteProcedure(ByVal enumFuncCaller As FunctionCaller, Optional ByVal IsUseExec As Boolean = True)
    ' ***
    ' A procedure that allows to call the procedures written as enums.
    ' It allows to call function with usage of ExecuteProcedure or not.
    ' See more: Functions.ExecuteProcedure.
    ' ***
    
    ' Accepts:
    '   - enumFunctionCaller [enum] - the code of the formula to be called
    '   - IsUseExcec [bool, optional] - True means that the function shall be called via Functions.ExecuteProcedure function,
    '                                   otherwise it will be called by application run method.
        
    ' Returns:
    '   - None
    
    Dim strFunctionName As String
    
    strFunctionName = GetFunctionName(enumFuncCaller)
    If IsUseExec Then
        Call Functions.ExecuteProcedure(strFunctionName)
    Else
        Call Application.Run(strFunctionName)
        
    End If
End Sub

Public Sub UpdateScrollBar(ByVal enumScrollBar As ScrollBar, ByVal enumFunction As FunctionCaller)
    ' *** This function is responsible for updating the scrollbar in the current page ***
    
    ' Accepts:
    '   - enumScrollBar [enum] - the scrollbar number that shall be updated
    '   - enumFunction [enum] - the name of the function that would be bound to a scrollbar.
    ' Returns:
    '   - None
    
    ' EDIT: Some parameters are commented out due to HORRIBLE performance impact of this operation.
    Dim shpScrollBar As Shape
    
    Set shpScrollBar = ActiveSheet.Shapes(GetScrollBarName(enumScrollBar))
    
    Call SetScrollBarCurrentValue(enumScrollBar, 0)
    With shpScrollBar
        With .ControlFormat
            .Value = GetScrollBarCurrentValue(enumScrollBar)
            .Max = GetScrollBarMax(enumScrollBar)
        End With
        
        .OnAction = GetFunctionName(enumFunction)
    End With
    
End Sub
' - END UTILS -

' - WORKBOOK MODES -

'Private Sub SelectWorkbookMode(ByVal enumWbModeNum As WorkbookMode)
'    ' *** Handles changing the workbook settings according to user preferences ***
'
'    ' Accepts:
'    '   - intNum [integer] - the demanded global application settings
'    ' Returns:
'    '   - None
'
'    Select Case enumWbModeNum
'        Case Is = INTRO
'            Call SetIntroductionWorksheetsVisibility
'            Call SetDevelopmentMode
'
'        Case Is = CINEMATIC
'            Call SetNonIntroductionWorksheetsVisibility
'            Call SetCinematicMode
'
'        Case Is = NORMAL
'            Call SetNonIntroductionWorksheetsVisibility
'            Call SetDevelopmentMode
'
'        Case Else
'            Debug.Print "No such workbook mode!"
'
'    End Select
'
'End Sub

' Too unstable, basically does nothing from the POV of dashboard, but destroyes global application settings.

'Private Sub SetNonIntroductionWorksheetsVisibility()
'    ' *** This procedure is responsible for setting the visibility of the worksheets, when non introduction mode is set***
'
'    Dim arrDashboardPages() As Variant
'    Dim i As Integer
'
'    arrDashboardPages = Sheet1.GetDashboardPagesEnums()
'    Worksheets(Getters_Setters.GetWorksheetsName(Backend)).Visible = xlVeryHidden
'
'    For i = LBound(arrDashboardPages) To UBound(arrDashboardPages)
'        Worksheets(Getters_Setters.GetWorksheetsName(arrDashboardPages(i))).Visible = xlSheetVisible
'    Next i
'
'    Worksheets(Getters_Setters.GetWorksheetsName(Introduction)).Visible = xlVeryHidden
'    Worksheets(Getters_Setters.GetWorksheetsName(StartingPage)).Select
'
'End Sub
'
'Private Sub SetIntroductionWorksheetsVisibility()
'    ' *** This procedure sets the worksheet visibility as it was the introduction mode selected ***
'
'    Dim arrDashboardPages() As Variant
'    Dim i As Integer
'
'    arrDashboardPages = Sheet1.GetDashboardPagesEnums()
'    Worksheets(Getters_Setters.GetWorksheetsName(Introduction)).Visible = xlSheetVisible
'    Worksheets(Getters_Setters.GetWorksheetsName(Backend)).Visible = xlVeryHidden
'
'    For i = LBound(arrDashboardPages) To UBound(arrDashboardPages)
'        Worksheets(Getters_Setters.GetWorksheetsName(arrDashboardPages(i))).Visible = xlSheetVeryHidden
'    Next i
'
'    Worksheets(Getters_Setters.GetWorksheetsName(Introduction)).Select
'
'End Sub

' Project is too unstable and annyoing to watch.
' Idea abandoned.
'Private Sub SetDevelopmentMode()
'    ' *** Sets the current worksheet to the development mode. ***
'
'    With Application
'        ' .DisplayFullScreen = False
'        ' .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
'        .DisplayFormulaBar = True
'    End With
'
'    With ActiveWindow
'        .DisplayHorizontalScrollBar = True
'        .DisplayVerticalScrollBar = True
'        .DisplayHeadings = True
'    End With
'
'    With Application.Windows(1)
'        .WindowState = xlMaximized
'        .DisplayGridlines = True
'        .DisplayWorkbookTabs = True
'        .DisplayHorizontalScrollBar = True
'        .DisplayVerticalScrollBar = True
'        .DisplayHeadings = True
'        .DisplayOutline = True
'    End With
'
'End Sub

'Public Sub SetCinematicMode(Optional ByVal varProxy As Variant = 0)
'    ' *** Sets the current worksheet to the cinematic mode. ***
'
'    With Application
'        ' .DisplayFullScreen = True
'        ' .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
'        .DisplayFormulaBar = False
'    End With
'
'    With ActiveWindow
'        .DisplayHorizontalScrollBar = False
'        .DisplayVerticalScrollBar = False
'        .DisplayHeadings = False
'    End With
'
'    With Application.Windows(1)
'        .WindowState = xlMaximized
'        .DisplayGridlines = False
'        .DisplayWorkbookTabs = False
'        .DisplayHorizontalScrollBar = False
'        .DisplayVerticalScrollBar = False
'        .DisplayHeadings = False
'        .DisplayOutline = False
'    End With
'
'End Sub

' - END WORKBOOK MODE -

' - REFRESH DATA SOURCE -

Private Sub HandleRefreshingContent()
    ' ***
    ' Main procedure responsible for refreshing the content of the data sources.
    ' Entry point for handling the operation.
    ' ***
    
    Dim intAnswer As Integer
    
    intAnswer = MsgBox(Texts.backend_RefreshInputBox, vbYesNo + vbQuestion, "Question")
    
    If intAnswer = vbYes Then
        Call Sheet1.ChangeGlobalDateByValFunction(0)
        Call RefreshDataSource
        Call Sheet1.ChangeGlobalDateByValFunction(Getters_Setters.GetBackendValue(LastPossibleDateAsInteger) - 1)
        Call Functions.SetCurrentSlicer(fatesInternalDate, Getters_Setters.GetBackendValue(CurrentDateAsInteger))
    End If
    
End Sub

Private Sub RefreshDataSource()
    ' *** Refreshes all tables in the project ***
    
    ThisWorkbook.RefreshAll
End Sub

' - END REFRESH DATA SOURCE -
' - Crash Preventer.
Private Sub RemoveAllValidationListsFromProject()
    ' ***
    ' Removes all validation lists from the project in order to prevent random crashes.
    
    ' Implementation: https://www.mrexcel.com/board/threads/vba-to-clear-data-validation-from-all-sheets.1158157/
    ' ***
    
    Dim wsSheetObject As Worksheet
    
    For Each wsSheetObject In Worksheets
        On Error Resume Next
        wsSheetObject.UsedRange.SpecialCells(xlCellTypeAllValidation).Validation.Delete
        On Error GoTo 0
    Next wsSheetObject
    
End Sub

' ------------------------------------------------------------------------------------------------------
' -------------------------------------- FUNCTIONS && CALLS --------------------------------------------
' ------------------------------------------------------------------------------------------------------

Public Sub RefreshDataSourceFunction(Optional ByVal varProxy As Variant = 0)
    ' *** Function that is responsible for refreshing the data source in the project ***
    
    Call HandleRefreshingContent
End Sub

Public Sub ChangeWorkbookModeFunction(Optional ByVal varProxy As Variant = 0)
    ' *** Function that is responsible for changing the workbook mode in the project ***
    
    Dim intNum As Integer
    
    intNum = CInt(Right(Application.Caller, 1))
    Worksheets(Getters_Setters.GetWorksheetsName(Introduction)).Select
    
End Sub

Public Sub RefreshDataSourceCall(Optional ByVal varProxy As Variant = 0)
    ' *** Function responsible for refreshing the data in the project ***
    
    Call ExecuteProcedure(WbFuncRefreshSource)
End Sub

Public Sub ChangeWorkbookModeCall(Optional ByVal varProxy As Variant = 0)
    
    Call ExecuteProcedure(WbfuncChangeWorkbookMode)
End Sub

' ------------------------------------------------------------------------------------------------------
' ------------------------------------------ WORKBOOK CALLBACKS ----------------------------------------
' ------------------------------------------------------------------------------------------------------

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not Cancel Then
        Call RemoveAllValidationListsFromProject
    End If
End Sub

Private Sub Workbook_Open()
    Worksheets(Getters_Setters.GetWorksheetsName(Introduction)).Select
    
    With Application.Windows(1)
        .WindowState = xlMaximized
        .DisplayGridlines = False
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayHeadings = False
        .DisplayOutline = False
    End With
    
End Sub
