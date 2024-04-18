Attribute VB_Name = "Getters_Setters"
Option Private Module
Option Explicit
' --------------------------------------------
' Project: Attack on Europe 2022
' Made by: Matthew Borcz
' GH Page: https://github.com/MBrcz
' --------------------------------------------

' ***
' This module stores the implementation of getters and setters for the commonly used in whole project parameters.
' ***

' ------------------------------------------------------------------------------
' ----------------------- PUBLIC FUNCTIONS / SUBS ------------------------------
' ------------------------------------------------------------------------------
' In this section I am not going to write comments.
' All of the function might be summarized to Get Param X of Y or Set Param X to value Z of Y
' Anyway, everyone knows how does "getters / setters" work, right?

' -- Worksheets --
Public Function GetWorksheetsName(ByVal enumProjectSheets As ProjectSheets) As String
    GetWorksheetsName = GetWorksheetParameter(enumProjectSheets, SheetName)
End Function
' -- END WORKSHEETS --

' -- SLICERS --
Public Function GetSlicerName(ByVal enumSlicerCall As SlicerCall) As String
    GetSlicerName = GetSlicerParameter(enumSlicerCall, SlicerCallParam.Name) ' For some mystical reason, lone Name throws error.
End Function

Public Function GetSlicerTable(ByVal enumSlicerCall As SlicerCall) As String
    GetSlicerTable = GetSlicerParameter(enumSlicerCall, SlicerCallParam.Table)
End Function

Public Function GetSlicerColumn(ByVal enumSlicerCall As SlicerCall) As String
    GetSlicerColumn = GetSlicerParameter(enumSlicerCall, SlicerCallParam.Column)
End Function

Public Function GetSlicerCurrentValue(ByVal enumSlicerCall As SlicerCall) As String
    GetSlicerCurrentValue = GetSlicerParameter(enumSlicerCall, CurrentValue)
End Function

Public Sub SetSlicerCurrentValue(ByVal enumSlicerCall As SlicerCall, strValue As String)
    Call SetSlicerParameter(enumSlicerCall, CurrentValue, strValue)
End Sub

Public Function GetSlicerSignature(ByVal enumSlicerCall As SlicerCall) As String
    GetSlicerSignature = GetSlicerParameter(enumSlicerCall, SignatureName)
End Function
' -- END SLICERS --

' -- SCROLLBARS --
Public Function GetScrollBarCurrentValue(ByVal enumScrollBar As ScrollBar) As String
    GetScrollBarCurrentValue = GetScrollbarParameter(enumScrollBar, Current)
End Function

Public Sub SetScrollBarCurrentValue(ByVal enumScrollBar As ScrollBar, intValue As String)
    Call SetScrollBarParameter(enumScrollBar, Current, intValue)
End Sub

Public Function GetScrollBarMin(ByVal enumScrollBar As ScrollBar) As String
    GetScrollBarMin = GetScrollbarParameter(enumScrollBar, Min)
End Function

Public Function GetScrollBarMax(ByVal enumScrollBar As ScrollBar) As String
    GetScrollBarMax = GetScrollbarParameter(enumScrollBar, Max)
End Function

Public Function GetScrollBarToShow(ByVal enumScrollBar As ScrollBar) As String
    GetScrollBarToShow = GetScrollbarParameter(enumScrollBar, ToShow)
End Function

Public Function GetScrollBarIncrChange(ByVal enumScrollBar As ScrollBar) As String
    GetScrollBarIncrChange = GetScrollbarParameter(enumScrollBar, IncrChange)
End Function

Public Function GetScrollBarName(ByVal enumScrollBar As ScrollBar) As String
    GetScrollBarName = GetScrollbarParameter(enumScrollBar, ScrName)
End Function

' -- END SCROLLBARS --

' -- OTHER PARAM --
Public Function GetOtherParam(ByVal enumOtherParamNum As OtherParamNum) As String
    GetOtherParam = GetOtherParameter(enumOtherParamNum, ParamValue)
End Function
' -- END OTHER PARAM --

' -- FAMILY OBJECT --
' -- BOUND TO MILITARY FAMILY FAMILIES SHEET!
Public Function GetFamilyImagePosition(ByVal enumFamilyObject As FamilyObject) As String
    GetFamilyImagePosition = GetFamilyObjectParameter(enumFamilyObject, ImageRange)
End Function

Public Function GetFamilyTextBoxRange(ByVal enumFamilyObject As FamilyObject) As String
    GetFamilyTextBoxRange = GetFamilyObjectParameter(enumFamilyObject, TextBoxRange)
End Function

Public Function GetFamilyObjectDescription(ByVal enumFamilyObject As FamilyObject) As String
    GetFamilyObjectDescription = GetFamilyObjectParameter(enumFamilyObject, Description)
End Function

Public Function GetFamilyImageUrl(ByVal enumFamilyObject As FamilyObject) As String
    GetFamilyImageUrl = GetFamilyObjectParameter(enumFamilyObject, ImageURLPath)
End Function
' -- END FAMILY OBJECT

' - MILITARY OVERVIEW BRANCH -
' - BOUND TO MILITARY OVERWIEV SHEET!
Public Function GetMilitaryOverwievSlicerValue(ByVal enumMilOverwievBranch As MilOverwievBranch) As String
    GetMilitaryOverwievSlicerValue = GetMilitaryOverwievParameter(enumMilOverwievBranch, MilValue)
End Function
' - END MILITARY OVERVIEW BRANCH -

Public Function GetItemSearcherRange(ByVal enumItemBoxes As ItemBoxes) As String
    GetItemSearcherRange = GetItemSearcherParameter(enumItemBoxes, SlicerBoxRange)
End Function

Public Function GetItemSearcherContentRangeName(ByVal enumItemBoxes As ItemBoxes) As String
    GetItemSearcherContentRangeName = GetItemSearcherParameter(enumItemBoxes, SlicerSourceRangeName)
End Function

Public Function GetItemSearcherSlicerToApply(ByVal enumItemBoxes As ItemBoxes) As String
    GetItemSearcherSlicerToApply = GetItemSearcherParameter(enumItemBoxes, SlicerToApply)
End Function

Public Function GetItemSearcherBookmarkParam(ByVal enumItemBoxes As ItemBoxes, ByVal enumItemSearcherBookmark As ItemSearcherBookmark) As String
    ' *** Please, do not do this at home. Dangerous coding ahead. ***
    GetItemSearcherBookmarkParam = GetItemSearcherParameter(enumItemBoxes, CInt(enumItemSearcherBookmark))
End Function
' - ITEMS SEARCHER END-

' - ITEMS OTHER PARAM -
Public Function GetItemOtherValue(ByVal enumItemOther As ItemOther) As String
    GetItemOtherValue = GetItemOtherParameter(enumItemOther, ItemOtherValue)
End Function
' - ITEMS OTHER PARAM END -

' - BACKEND -
Public Function GetBackendValue(ByVal enumBackend As Backend) As String
    GetBackendValue = GetBackendParameter(enumBackend, BackendValue)
End Function
' - END BACKEND -

' - FUNCTION CALLER -
Public Function GetFunctionName(ByVal enumFunctionCaller As FunctionCaller) As String
    GetFunctionName = GetFunctionCallerParameter(enumFunctionCaller, FuncValue)
End Function

' - END FUNCTION CALLER -

' -------------------------------------------------------------------------------
' ----------------------------- PRIVATE FUNCTIONS -------------------------------
' -------------------------------------------------------------------------------

Private Function GetBackendParameter(ByVal enumBackend As Backend, ByVal enumBackendParam As BackendParam) As String
    ' *** Gets and returns the parameter from the Backend parameter table. ***
    
    ' Accepts:
    '   - enumBackend [enum] - the number of item assigned to a parameter
    '   - enumBackendParam [enum] - the parameter name assigned to a item
    ' Returns:
    '   - string - parameter value
    
    Dim rngBackendParameterTableStart As Range
    Set rngBackendParameterTableStart = GetTheStartTable(BackendParamaters)
    
    GetBackendParameter = FindKeyInTable(rngBackendParameterTableStart, enumBackend).Offset(0, enumBackendParam)
End Function

Private Function GetItemOtherParameter(ByVal enumItemOther As ItemOther, ByVal enumItemOtherParam As ItemOtherParam) As String
    ' *** Gets and returns the paramter from item other table parameter ***
    
    ' Accepts:
    '   - enumItemOther [enum] - the number of assigned to the parameter to get
    '   - enumItemOtherParam [enum] - the parameter that will be returned
    ' Retuns:
    '   - string - parameter value
    
    Dim rngItemOtherParameterTableStart As Range
    Set rngItemOtherParameterTableStart = GetTheStartTable(ItemsOtherParameters)
    
    GetItemOtherParameter = FindKeyInTable(rngItemOtherParameterTableStart, enumItemOther).Offset(0, enumItemOtherParam)
End Function

Private Function GetItemSearcherParameter(ByVal enumItemBoxes As ItemBoxes, _
                                          ByVal enumItemBoxesParam As ItemBoxesParam) As String
    ' *** Gets and returns the parameter of a chosen item searcher textbox parameter. ***
    
    ' Accepts:
    '   - enumItemBoxes [enum] - the number of a item searcher textbox that shall be checked.
    '   - enumItemBoxesParam [enum] - the parameter value that will be returned from this operation.
    
    Dim rngItemSearcherParameterTableStart As Range
    Set rngItemSearcherParameterTableStart = GetTheStartTable(ItemsSearcherParameters)
    
    GetItemSearcherParameter = FindKeyInTable(rngItemSearcherParameterTableStart, enumItemBoxes).Offset(0, enumItemBoxesParam)
End Function

Private Function GetFunctionCallerParameter(ByVal enumFunctionCaller As FunctionCaller, enumFunctionCallerParam As FunctionCallerParam) As String
    ' *** Gets and returns the parameter of a passed function from the project ***
    
    ' Accepts:
    '   - enumFunctionCaller [enum] - the name of the function as enum
    '   - enumFunctionParameter [enum] - the name of the parameter in question
    
    ' Returns:
    '   - string - asked parameter value
    
    Dim rngFunctionCallerObjectTableStart As Range
    Set rngFunctionCallerObjectTableStart = GetTheStartTable(FunctionCaller)
    
    GetFunctionCallerParameter = FindKeyInTable(rngFunctionCallerObjectTableStart, enumFunctionCaller).Offset(0, enumFunctionCallerParam)
    
End Function

Private Function GetMilitaryOverwievParameter(ByVal enumMilOverwievBranch As MilOverwievBranch, _
                                              ByVal enumMilOverwievBranchParam As MilOverwievBranchParam) As String
    ' *** Gets and returns the parameter related to the military overwiev. ***
    
    ' Accepts:
    '   - enumMilOverwievBranch [enum] - the branch num that will be returned.
    '   - enumMilOverwievBranchParam [enum] - the branch parameter name that will be returned.
    
    ' Returns:
    '   - string - asked parameter value
    
    Dim rngMilitaryOverwievObjectTableStart As Range
    Set rngMilitaryOverwievObjectTableStart = GetTheStartTable(MilitaryBranchParameter)
    
    GetMilitaryOverwievParameter = FindKeyInTable(rngMilitaryOverwievObjectTableStart, enumMilOverwievBranch).Offset(0, enumMilOverwievBranchParam)

End Function

Private Function GetFamilyObjectParameter(ByVal enumFamilyObject As FamilyObject, ByVal enumFamilyObjectParam As FamilyObjectParam) As String
    ' *** Gets and returns the parameter related to the machinations of family objects ***
    
    ' Accepts:
    '   - enumFamilyObject [enum] - the number of the family object in question
    '   - enumFamiliyObjectParam [enum] - the name of the parameter to be retruned
    
    ' Returns:
    '   - None
    
    Dim rngFamilyObjectTableStart As Range
    Set rngFamilyObjectTableStart = GetTheStartTable(FamilyParameters)
    
    GetFamilyObjectParameter = FindKeyInTable(rngFamilyObjectTableStart, enumFamilyObject).Offset(0, enumFamilyObjectParam).Value
    
End Function

Private Function GetOtherParameter(ByVal enumOtherParamNum As OtherParamNum, ByVal enumOtherParamName As OtherParamName) As String
    ' *** Gets and returns other parameter used in the project. ***
    
    ' Accepts:
    '   - enumOtherParamNum [enum] - the number of the asked parameter
    '   - enumOtherParamName [enum] - the name of the returned parameter.
    
    ' Returns:
    '   - string - asked parameter value
    
    Dim rngOtherParamTableStart As Range
    Set rngOtherParamTableStart = GetTheStartTable(OtherParameters)

    GetOtherParameter = FindKeyInTable(rngOtherParamTableStart, enumOtherParamNum).Offset(0, enumOtherParamName).Value
    
End Function

Private Function GetWorksheetParameter(ByVal enumProjectSheets As ProjectSheets, ByVal enumProjectSheetsParam As ProjectSheetsParam) As String
    ' *** Gets the demanded parameter of a chosen by user worksheet ***
    
    ' Accepts:
    '   - enumProjectSheets [enum] - the number of the asked sheet
    '   - enumProjectSheetsParam [enum] - the number of the parameter that will be returned.
    
    ' Returns:
    '   - string - the asked parameter of the worksheet.
    
    Dim rngWorksheetTableStart As Range
    Set rngWorksheetTableStart = GetTheStartTable(WorksheetNames)

    GetWorksheetParameter = FindKeyInTable(rngWorksheetTableStart, enumProjectSheets).Offset(0, enumProjectSheetsParam).Value
    
End Function

Private Function GetSlicerParameter(ByVal enumSlicerCall As SlicerCall, ByVal enumSlicerParam As SlicerCallParam) As String
    ' *** Gets and returns the demanded by user parameter related to the slicers. ***
    
    ' Accepts:
    '   - enumSlicerCall [enum] - the number of slicer that will be called
    '   - enumSlicerParam [enum] - the number of the parameter that will be returned.
    ' Returns:
    '   - String - the demanded parameter.
    
    Dim rngSlicerTableStart As Range
    
    Set rngSlicerTableStart = GetTheStartTable(SlicerCall)
    GetSlicerParameter = FindKeyInTable(rngSlicerTableStart, enumSlicerCall).Offset(0, enumSlicerParam).Value

End Function

Private Sub SetSlicerParameter(ByVal enumSlicerCall As SlicerCall, ByVal enumSlicerParam As SlicerCallParam, ByVal strValue As String)
    ' *** Sets the parameter of the slicer to the new value.***
    
    ' Accepts:
    '   - enumSlicerCall [enum] - the slicer that will be called
    '   - enumSlicerParam [enum] - the parameter that will be changed
    '   - strValue [string] - the new value that will be set
    
    ' Returns:
    '   - None
    
    Dim rngSlicerTableStart As Range
    
    Set rngSlicerTableStart = GetTheStartTable(SlicerCall)
    FindKeyInTable(rngSlicerTableStart, enumSlicerCall).Offset(0, enumSlicerParam).Value = strValue
    
End Sub

Private Sub SetScrollBarParameter(ByVal enumScrollBar As ScrollBar, ByVal enumScrollbarParam As ScrollBarParam, ByVal strValue As String)
    ' *** Sets the parameter of a demanded scrollbar to value. ***
    
    ' Accepts:
    '   - enumScrollBar [enum] - the scrollbar that will be checked.
    '   - enumScrollBarParam [enum] = the parameter of the scrollbar that will be changed
    '   - strValue [str] - the value that will be set
    
    ' Returns:
    '   - None
    
    Dim rngScrollBarTableStart As Range
    Set rngScrollBarTableStart = GetTheStartTable(WorkbookScrollbars)
    
    FindKeyInTable(rngScrollBarTableStart, enumScrollBar).Offset(0, enumScrollbarParam).Value = strValue
    
End Sub

Private Function GetScrollbarParameter(ByVal enumScrollBar As ScrollBar, ByVal enumScrollbarParam As ScrollBarParam) As String
    ' *** Gets and returns the value of the demanded scrollbar parameter.***
    
    ' Accepts:
    '   - enumScrollBar [enum] - the number of the scrollbar that shall probide parameters
    '   - enumScrollBarParam [num] - the parameter of the scrollbar that shall be returned
    
    ' Returns:
    '   - string - the demanded scrollbar parameter.
    
    Dim rngScrollBarTableStart As Range
    
    Set rngScrollBarTableStart = GetTheStartTable(WorkbookScrollbars)
    GetScrollbarParameter = FindKeyInTable(rngScrollBarTableStart, enumScrollBar).Offset(0, enumScrollbarParam).Value

End Function

' ---------------------------------------------------
' ------------------------ CORE ---------------------
' ---------------------------------------------------

Private Function FindKeyInTable(ByVal rngStartRange As Range, ByVal intKeyValue As Integer) As Range
    ' *** Finds the key value in the asked table ***
    
    ' Accepts:
    '   - rngStartRange [range] - the range from which, the looping shall start
    '   - intKeyValue [integer] - the value of the item that shall be searched.
    ' Returns:
    '   - Range - the position of the key in the asked table.
        
    Dim i As Integer

    i = 0
    Do While rngStartRange.Offset(i, 0).Value <> intKeyValue
        If rngStartRange.Offset(i, 0).Value = "" Then
            MsgBox "Cannot find such key!", vbCritical
            Exit Function
        End If

        i = i + 1
    Loop

    Set FindKeyInTable = rngStartRange.Offset(i, 0)

End Function

Private Function GetTheStartTable(ByRef enumStartTables As StartTables) As Range
    ' *** Gets the range of the start table for the known table.***
    
    ' Accepts:
    '   - enumStartTables [enum] - the number of table that start position is going to be returned
    ' Returns:
    '   - Range = the new position of the table.
    
    Dim rngStartTable As Range
    Dim rngStartTableKey As Range
    
    Set rngStartTable = Sheet1.Names("_backend_start_table").RefersToRange
    Set rngStartTableKey = FindKeyInTable(Range(rngStartTable), enumStartTables)
    
    Set GetTheStartTable = Range(rngStartTableKey.Offset(0, Reference))
    
End Function

