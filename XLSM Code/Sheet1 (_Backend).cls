VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
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

' --------------------------------------------
' -------------- DATES HANDLING --------------
' --------------------------------------------

' INPUT BOX
' TEXTS
Private Function GetDatesToChooseAsText() As String
    ' *** Formats the dates as text to the text form. Usable in Input Box question. ***
    
    ' Accepts:
    '   - None
    ' Returns:
    '   - string - the text value of a dates with appropriate numbers and "\n" mark.
    
    Dim arrDatesAsText() As Variant
    Dim strDates As String
    
    Dim i As Integer
    
    arrDatesAsText = ConvertRangeToArray(Range(GetBackendValue(RangeDatesAsDates)), eBottom, False)
    
    i = 0
    strDates = ""
    For i = LBound(arrDatesAsText) To UBound(arrDatesAsText)
        strDates = strDates & i + 1 & ": " & arrDatesAsText(i) & Chr(10)
    Next i
    
    GetDatesToChooseAsText = strDates
End Function

' - END TEXT

' - FUNCTIONS

Private Sub ChangeDatesInProject()
    ' *** This procedure is responsible for changing the current and previous dates values in the project. ***
    
    Dim intUserInput As Integer
    Dim arrDatesToChoose As Variant
    
    intUserInput = GetUserInput()
    
    If intUserInput <> 0 Then
        Call SwitchContentOfDateSlicers(intUserInput - 1)
    End If

End Sub

Private Function GetUserInput() As Integer
    ' *** This function is responsible for getting user input in response to date changing operation ***
    
    Dim intCountPossibleChoices As Integer
    Dim strMsgInfo As String
    Dim varAnswer As Variant

    intCountPossibleChoices = UBound(ConvertRangeToArray(Range(GetBackendValue(RangeDatesAsNums)), eBottom, False)) + 1
    strMsgInfo = Texts.backend_dateInput(GetDatesToChooseAsText) & GetBackendValue(CurrentDateAsText)
    
    varAnswer = ""
    Do While True
        varAnswer = InputBox(strMsgInfo, "Type a Num")
        
        If varAnswer = "" Then
            GetUserInput = 0
            Exit Do
        
        ElseIf (0 <= varAnswer And varAnswer <= intCountPossibleChoices) Then
            GetUserInput = varAnswer
            Exit Do
        End If
    
        MsgBox Texts.backend_dateError, vbCritical + vbOKOnly
    Loop
    
End Function

Private Sub SwitchContentOfDateSlicers(ByVal intDateNum As Integer)
    ' ***
    ' This function is responsible for changing the content of both date slicers in project.
    ' ***
    
    ' Accepts:
    '   - intDateNum [integer] - the number of which date shall be chosen. It starts from 0, not 1 (IMPORTANT!)
    ' Returns:
    '   - None
    
    Dim arrDatesToChoose As Variant
    
    arrDatesToChoose = ConvertRangeToArray(Range(GetBackendValue(RangeDatesAsNums)), eBottom, False)
    Call SetCurrentSlicer(backCurrentDate, CStr(arrDatesToChoose(intDateNum)))
    
    If intDateNum - 1 <= 0 Then
        Call SetCurrentSlicer(backPreviousDate, CStr(arrDatesToChoose(0)))
    Else
        Call SetCurrentSlicer(backPreviousDate, CStr(arrDatesToChoose(intDateNum - 1)))
    End If
    
End Sub

' -------------------------------------------------------
' ------------------- WORKSHEET MOVEMENET ---------------
' -------------------------------------------------------

Public Function GetDashboardPagesEnums() As Variant()
     ' *** This function stores all dashboard pages names in correct order ***
     
     GetDashboardPagesEnums = Array(StartingPage, MilitaryOverwiev, MilitaryFamilies, ItemsSearcher, ItemsFates)
End Function

Private Sub MoveToWorksheet(ByVal intDestinationSheetNum As Integer)
    ' *** This procedure is responsible for selecting a worksheet in the project ***
    
    ' Accepts:
    '   - intDestinationSheetNum [integer] - the number from backend which the sheet will be moved to.
    ' Returns:
    '   - None
    
    Dim strDestinationSheetName As String
    
    strDestinationSheetName = Getters_Setters.GetWorksheetsName(intDestinationSheetNum)
    Worksheets(strDestinationSheetName).Select
    
End Sub

Private Sub MoveToWorksheetRelatively(ByVal boolIsNextSheet As Boolean)
    ' *** Moves to the next worksheet relatively from the current one ***
    
    ' Accepts:
    '   - boolIsNextSheet [boolean] - True means that script is going to the next worksheet, otherwise to previous.
    
    Dim strIteratedWorksheetName As String
    Dim arrWorksheetNames As Variant
    Dim i As Integer
    
    arrWorksheetNames = GetDashboardPagesEnums()
    
    On Error Resume Next
    For i = LBound(arrWorksheetNames) To UBound(arrWorksheetNames)
        strIteratedWorksheetName = Getters_Setters.GetWorksheetsName(arrWorksheetNames(i))
                
        If strIteratedWorksheetName = ActiveSheet.Name And boolIsNextSheet Then
            Call MoveToWorksheet(arrWorksheetNames(i) + 100)
            Exit Sub

        ElseIf strIteratedWorksheetName = ActiveSheet.Name And Not boolIsNextSheet Then
            Call MoveToWorksheet(arrWorksheetNames(i) - 100)
            Exit Sub
        End If
    Next i
  
End Sub

' -------------------------------------------------------
' ---------------------- FLAGS --------------------------
' -------------------------------------------------------

Private Sub GoToSource(ByVal intSourceNum As Integer)
    ' *** Checks the source of the data for current date ***
    
    ThisWorkbook.FollowHyperlink Getters_Setters.GetBackendValue(intSourceNum)
End Sub


' -------------------------------------------------------
' ------------------- FUNCTIONS -------------------------
' -------------------------------------------------------

Public Sub UpdateControlPanelObjects(Optional ByVal varProxy As Variant = 0)
    ' *** This function is responsible for invoking the function into an backend objects ***
    
    Dim arrObjectNames() As Variant
    Dim arrFunctionsNames() As Variant
    Dim i As Integer
    
    arrObjectNames = Array(ChangeDateButtonSignature, ArrowMoveSignature, FlagSignature, BackendLogo)
    arrFunctionsNames = Array(S1CallChangeGlobalDate, S1CallMoveToSheetViaArrow, S1CallCheckTheSource, WbCallChangeWorkbookMode)
    
    For i = LBound(arrObjectNames) To UBound(arrObjectNames)
        Call Functions.InvokeFunctionWithinAnImage( _
                       strImageName:=Getters_Setters.GetBackendValue(arrObjectNames(i)), _
                       strCommandName:=Getters_Setters.GetFunctionName(arrFunctionsNames(i)) _
        )
    Next i
    
    ' This is kinda desperate, but who cares?
   '  Call ThisWorkbook.SetCinematicMode
    
End Sub

Public Sub ChangeGlobalDateFunction(Optional ByVal varProxy As Variant = 0)
    ' *** Changes the global date function according to the user input ***
    Call ChangeDatesInProject
End Sub

Public Sub ChangeGlobalDateByValFunction(ByVal intDateNum As Integer)
    ' ***
    ' Changes the global date of the project chosen by an user.
    ' This function does not take into account GUI elements.
    ' ***
    
    ' Accepts:
    '   - intDateNum [integer] - the number of the date that shall be set
    ' Returns:
    '   - None
    
    Call SwitchContentOfDateSlicers(intDateNum)
End Sub

' Not not implemented, due to annoyingness it can cause.
Public Sub MoveToNextSheetFunction(Optional ByVal varProxy As Variant = 0)
    ' *** Moves to the next worksheet starting from current one. ***
    
    Call MoveToWorksheetRelatively(True)
End Sub

' Not not implemented, due to annoyingness it can cause.
Public Sub MoveToPreviousSheetFunction(Optional ByVal varProxy As Variant)
    ' *** Moves to the previous worksheet starting from current one. ***
    
    Call MoveToWorksheetRelatively(False)
End Sub

Public Sub MoveToSheetViaArrowFunction(Optional ByVal varProxy As Variant = 0)
    ' *** This function is responsible for moving to the other workhseets in project ***
    
    Dim intArrNum As Integer
    intArrNum = CInt(Right(Application.Caller, 1) & "00")
    
    Call MoveToWorksheet(intArrNum)
End Sub

Public Sub CheckTheSourceFunction(Optional ByVal varProxy As Variant = 0)
    ' *** This function is responsible for moving to the source of the data for the project. ***
    
    Dim intSourceNum As Integer
    intSourceNum = CInt(Right(Application.Caller, 1) & "00")
    
    Call GoToSource(intSourceNum)
End Sub

' ---------------------------------------------------------------------------
' ---------------------------- CALL -----------------------------------------
' ---------------------------------------------------------------------------

Public Sub MoveToSheetViaArrowCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S1FuncMoveToSheetViaArrow)
End Sub

Public Sub ChangeGlobalDateCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S1FuncChangeGlobalDate)
End Sub

Public Sub CheckTheSourceCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S1FuncCheckTheSource)
End Sub

