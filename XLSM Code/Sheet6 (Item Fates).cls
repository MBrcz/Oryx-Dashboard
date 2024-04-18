VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
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

Private Function GetTextBoxSlicersCodes() As Variant
    ' *** Gets the codes of all the slicers that have their input boxes in page. ***
    
    GetTextBoxSlicersCodes = Array(ItemFatesBranch, ItemFatesEquipmentType, ItemFatesFamilyNames, ItemFatesFullName, _
                                   ItemFatesCountryFullName)
End Function

' - PORTRAIR SHENENIGANS -
Private Sub HandlePortrairOperations()
    ' *** This procedure is responsible for managing the portrairs operations in the page ***
    
    Dim strPortrairSingature As String
   
    strPortrairSingature = Getters_Setters.GetItemOtherValue(ItemFatesPortrairSignature)
    
    Call Functions.RemoveImagesWithSignature(strPortrairSingature)
    Call ChangePortrairImage(strPortrairSingature)
    Call ChangePortrairTextBox
    
End Sub

Private Sub ChangePortrairImage(ByVal strPortrairSignature As String)
    ' *** This procedure handles changing the portrair image depending on selected slicers ***
    
    ' Accepts:
    '   - strPortrairSignature [string] - the signature of the newly added portrair image
    ' Returns:
    '   - None
    
    Dim rngPortrairRange As Range
    Dim strRangeName As String
    Dim strPortrairPath As String
    
    Set rngPortrairRange = Range(GetItemOtherValue(ItemFatesPortrairRange))
    strPortrairPath = Getters_Setters.GetItemOtherValue(ItemFatesPortrairPath)
 
    Call Functions.PlaceImage( _
                   rngImageRange:=rngPortrairRange, _
                   strPath:=strPortrairPath, _
                   newImageName:=strPortrairSignature _
    )
                   
End Sub

Private Sub ChangePortrairTextBox()
    ' *** This procedure handles the operation of changing the text of portrair TextBox ***
    
    Dim shpTextBox As Shape
    
    Dim strPortrairTextBoxName As String
    Dim strTextToPass As String
    
    strPortrairTextBoxName = Getters_Setters.GetItemOtherValue(ItemFatesPortrairTextBoxName)
    strTextToPass = Getters_Setters.GetItemOtherValue(ItemFatesTextBoxContent)

    Set shpTextBox = ActiveSheet.Shapes(strPortrairTextBoxName)
    
    With shpTextBox
        With .TextFrame
            With .Characters
                .Text = strTextToPass
                With .Font
                    .Bold = msoTrue
                    .Size = 12
                    .Name = "Time New Roman"
                End With
            End With
            
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignCenter
        End With
    End With
    
End Sub

' - END PORTRAIR SHENENIGANS

' -------------------------------------------------------
' ------------------ TEXT BOXES -------------------------
' -------------------------------------------------------

Private Sub InjectBookmarksToTextBoxes(ByVal enumItemFatesBookmark As ItemSearcherBookmark)
    ' *** Places the bookmarks parameters in the text boxes in current module. ***
    
    ' Accepts:
    '   - enumItemFatesBookmark [enum] - the number of the bookmark parameter that shall be set.
    ' Returns:
    '   - None
        
    Dim arrTextBoxCodes() As Variant
    Dim varTextBoxCode As Variant
    Dim rngTextBox As Range

    arrTextBoxCodes = GetTextBoxSlicersCodes()

    For Each varTextBoxCode In arrTextBoxCodes
        Set rngTextBox = Range(Getters_Setters.GetItemSearcherRange(varTextBoxCode))
        rngTextBox.Value = Getters_Setters.GetItemSearcherBookmarkParam(varTextBoxCode, enumItemFatesBookmark)
    Next varTextBoxCode
    
End Sub

Private Sub InjectElementsToChooseToTextBoxes()
    ' ***
    ' As name implies it injects the lists with content to choose.
    ' They are automatically filtered
    ' ***
    
    Dim arrTextBoxCodes() As Variant
    Dim varTextBoxCode As Variant
    Dim rngTextBox As Range
    
    Dim strRangeName As String
    Dim strFormulaName As String
    
    arrTextBoxCodes = GetTextBoxSlicersCodes()
    
    For Each varTextBoxCode In arrTextBoxCodes
        Set rngTextBox = Application.Range(Getters_Setters.GetItemSearcherRange(varTextBoxCode))
        strRangeName = Getters_Setters.GetItemSearcherContentRangeName(varTextBoxCode)
        strFormulaName = "=INDIRECT(" & Chr(34) & strRangeName & Chr(34) & ")"

        Call Functions.InjectContentToValidationList(rngTextBox, strFormulaName)
    Next varTextBoxCode
    
End Sub

Private Sub ClearTheTextBoxes()
    ' *** Clears the content from the textboxes completely. ***
    
    Dim arrTextBoxCodes() As Variant
    Dim varTextBoxCode As Variant
    Dim rngTextBox As Range

    arrTextBoxCodes = GetTextBoxSlicersCodes()

    For Each varTextBoxCode In arrTextBoxCodes
        Set rngTextBox = Range(Getters_Setters.GetItemSearcherRange(varTextBoxCode))
        rngTextBox.Value = ""
    Next varTextBoxCode
    
End Sub

' -- CHECK SLICER CORECTNESS --
Private Function AreThoseSlicersPossibleToApply() As Boolean
    ' ***
    ' This function is responsible for checking if it is possible to apply the asked
    ' combination of slicers in the project without returning nothing.
    ' ***
    
    ' Accepts:
    '   - None
    ' Returns:
    '   - bool True if it is possible, otherwise False
    
    Dim arrSlicersToApply() As Variant
    
    arrSlicersToApply = GetAllSlicersToApply()
    
    If Functions.IsArrayEmpty(arrSlicersToApply) Then
        AreThoseSlicersPossibleToApply = True
        
    ElseIf IsSlicerCombinationInBaseTable(arrSlicersToApply) Then
        AreThoseSlicersPossibleToApply = True
    
    Else
        AreThoseSlicersPossibleToApply = False
    End If

    ' Debug.Print "Slicers possbile to pass: " & AreThoseSlicersPossibleToApply
End Function

Private Function GetAllSlicersToApply() As Variant()
    ' ***
    ' This function is responsible for getting all slicers that are going to be applied.
    ' Ignores all slicers that have value All.
    ' ***
    
    Dim arrTextBoxCodes() As Variant
    Dim arrOffsetSlicer() As Variant
    Dim strSlicerToApply As String
    Dim i As Integer
    
    arrTextBoxCodes = GetTextBoxSlicersCodes()

    For i = LBound(arrTextBoxCodes) To UBound(arrTextBoxCodes)
        strSlicerToApply = Getters_Setters.GetItemSearcherSlicerToApply(arrTextBoxCodes(i))
        
        If strSlicerToApply <> "All" Then
            ReDim Preserve arrOffsetSlicer(i)
            arrOffsetSlicer(i) = strSlicerToApply
        End If
    Next i
    
    GetAllSlicersToApply = arrOffsetSlicer
    
End Function

Private Function IsSlicerCombinationInBaseTable(ByRef arrOffsetSlicer() As Variant) As Boolean
    ' ***
    ' This function is responsible for looping for the slicer table and checking
    ' whether the combination of a passed slicers can be delivered.
    ' ***
    
    Dim rngTableNextPoint As Range
    Dim strTableName As String
    
    Dim i As Integer
    
    strTableName = Getters_Setters.GetItemOtherValue(ItemFatesCorrectnessTable)
    
    Set rngTableNextPoint = Application.Range(strTableName)
    
    Do While rngTableNextPoint.Offset(1, 0) <> ""
        Set rngTableNextPoint = rngTableNextPoint.Offset(1, 0)
        
        If IsMatchingRow(rngTableNextPoint, arrOffsetSlicer) Then
            IsSlicerCombinationInBaseTable = True
            Exit Function
        End If
        
        i = i + 1
    Loop
    
    IsSlicerCombinationInBaseTable = False

End Function

Private Function IsMatchingRow(ByVal rngFirstRange As Range, ByRef arrOffsetSlicers() As Variant) As Boolean
    ' *** Checks if current row contain parameters asked in slicers to pass. ***
    
    ' Accepts:
    '   - rngFirstRange [rng] - the range of first column in current row.
    '   - arrOffsetSlicers [arr] - the array that contains slicers that shall be applied
    ' Returns:
    '   - bool - True if it matches, otherwise False
    
    Dim i As Integer

    For i = LBound(arrOffsetSlicers) To UBound(arrOffsetSlicers)
        If rngFirstRange.Offset(0, i).Value <> arrOffsetSlicers(i) And _
            Not IsEmpty(arrOffsetSlicers(i)) Then
                IsMatchingRow = False
            Exit Function
         End If
    Next i

    IsMatchingRow = True
    
End Function

' -- END CHECK SLICER CORECTNESS --

' ------------------------------------------------------------
' ------------------------- SLICERS --------------------------
' ------------------------------------------------------------

Private Sub ClearTextBoxSlicers()
    ' ***
    ' Clears all textbox slicers.
    ' Does not work for Side of Country slicer.
    ' ***
        
    Dim arrTextBoxCodes() As Variant
    Dim varTextBoxCode As Variant

    Dim strAppliedSlicer As String
    
    arrTextBoxCodes = GetTextBoxSlicersCodes()

    For Each varTextBoxCode In arrTextBoxCodes
        strAppliedSlicer = Getters_Setters.GetSlicerCurrentValue(varTextBoxCode)
        
        If strAppliedSlicer <> "All" Then
            Call Functions.ClearCurrentSlicer(varTextBoxCode)
            ' Debug.Print "Cleared the slicer with code: " & varTextBoxCode & " || applied slicer " & strAppliedSlicer
        End If
    Next varTextBoxCode
    
End Sub

Private Sub SetTextBoxSlicers()
    ' ***
    ' Sets the Text Box slicers according to the content of the Text Boxes.
    ' Does not work for Side of Country slicer.
    ' ***
    
    Dim arrTextBoxCodes() As Variant
    Dim varTextBoxCode As Variant

    Dim strAppliedSlicer As String
    Dim strSlicerToApply As String

    arrTextBoxCodes = GetTextBoxSlicersCodes()

    For Each varTextBoxCode In arrTextBoxCodes
        strSlicerToApply = Getters_Setters.GetItemSearcherSlicerToApply(varTextBoxCode)
        strAppliedSlicer = Getters_Setters.GetSlicerCurrentValue(varTextBoxCode)

        If strSlicerToApply <> strAppliedSlicer Then
            Call Functions.SetCurrentSlicer(varTextBoxCode, strSlicerToApply)
        End If
    Next varTextBoxCode
        
End Sub

' - SIDE OF CONFLICT SLICER -
Public Sub InjectSideOfConflictSlicer(ByVal enumItemOther As ItemOther)
    ' *** This function is responsible for injecting / deinjecting the side of conflict slicer. ***
    
    ' Accepts:
    '   - enumOtherParamNum [enum] - the number of the slicer applied
    ' Returns:
    '   - None
    
    Dim strSlicerApplied As String
    Dim strSlicerToApply As String
    
    ' Tests if wrong argument has not been supplied.
    If enumItemOther <> ItemSideOfConflictRussia And enumItemOther <> ItemSideOfConflictUkraine Then
        MsgBox Texts.wrong_ArgumentError, vbCritical
        Exit Sub
    End If
    
    strSlicerApplied = Getters_Setters.GetSlicerCurrentValue(fatesSideOfConflict)
    strSlicerToApply = Getters_Setters.GetItemOtherValue(enumItemOther)
    
    Call ClearSideOfConflictSlicer
    If strSlicerApplied <> strSlicerToApply Then
        Call Functions.SetCurrentSlicer(fatesSideOfConflict, strSlicerToApply)
        Call Functions.SetTheLineInImages(Getters_Setters.GetItemOtherValue(ItemSideOfConflictSignature) & enumItemOther)
    End If
    
End Sub

Private Sub ClearSideOfConflictSlicer()
    ' *** This function is responsible for removing the side of conflict slicer in the current page. ***
     
    Call Functions.ClearTheLineInImages(Getters_Setters.GetItemOtherValue(ItemSideOfConflictSignature))
    Call Functions.ClearCurrentSlicer(fatesSideOfConflict)

End Sub
' - END SIDE OF CONFLICT -

' - OTHER -
Public Sub RefreshObjectsBinds()
    ' *** This function is responsible for refreshing the binds / text box sources. ***
    
    Dim shpObject As Shape
    
    Dim arrObjectNames() As Variant
    Dim arrFunctionsNames() As Variant
    Dim i As Integer
    
    arrObjectNames = Array(ItemFatesRestartFiltersSignature, ItemFatesSearchButtonSignature, _
                            ItemFatesBookmarkSignature, ItemSideOfConflictSignature)
    arrFunctionsNames = Array(S5CallRestartFitlers, S5CallSearchQuery, _
                             S5CallInvokeBookmarks, S5CallSetSideOfConflictSlicer)
    
    For i = LBound(arrObjectNames) To UBound(arrObjectNames)
        Call Functions.InvokeFunctionWithinAnImage( _
                       strImageName:=Getters_Setters.GetItemOtherValue(arrObjectNames(i)), _
                       strCommandName:=Getters_Setters.GetFunctionName(arrFunctionsNames(i)) _
        )
    Next i
    
    Call InjectElementsToChooseToTextBoxes
End Sub


' - END OTHER -
' --------------------------------------------------------------
' -------------------- FUNCTIONS -------------------------------
' --------------------------------------------------------------

Public Sub WorksheetOpenFunction(Optional ByVal varProxy As Variant = 0)
    ' *** This function is responsible for updating an objects durning selection sheets ***
    
    If ActiveSheet.Name = Getters_Setters.GetWorksheetsName(ItemsFates) Then
        Call RefreshObjectsBinds
        Call ThisWorkbook.ExecuteProcedure(FFuncZoomToVisibleCells, False)
        Call ThisWorkbook.ExecuteProcedure(S1FuncUpdateControlPanelObjects, False)
    End If
End Sub

Public Sub SearchQueryFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for handling searching the new query in the project
    ' ***
    
    If AreThoseSlicersPossibleToApply Then
        Call ClearTextBoxSlicers
        Call SetTextBoxSlicers
        Call HandlePortrairOperations
    Else
        MsgBox Texts.items_WrongQuery(), vbCritical + vbOKOnly, "Error"
    End If
End Sub

Public Sub RestartFiltersFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for restarting all the slicers applied in the page
    ' ***
    
    Call ClearSideOfConflictSlicer
    Call ClearTheTextBoxes
    Call ClearTextBoxSlicers
    Call HandlePortrairOperations
End Sub

Public Sub PlaceBookmarkFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for applying the filters in the project.
    ' ***

    Dim intNum As Integer

    intNum = CInt(Right(Application.Caller, 1))

    Call ClearTextBoxSlicers
    Call InjectBookmarksToTextBoxes(intNum)
    Call SetTextBoxSlicers
    Call HandlePortrairOperations
End Sub


Public Sub SetSideOfConflictSlicerFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for applying the side of conflicts slicers in the page.
    ' ***
    
    Dim intNum As Integer
    
     intNum = CInt(Right(Application.Caller, 1))
     Call InjectSideOfConflictSlicer(intNum)
End Sub

' ----------------------------------------------------------------
' ------------------------- CALLERS ------------------------------
' ----------------------------------------------------------------

Public Sub PlaceBookmarkCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S5FuncInvokeBookmarks)
End Sub

Public Sub SearchQueryCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S5FuncSearchQuery)
End Sub

Public Sub RestartFiltersCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S5FuncRestartFilters)
End Sub

Public Sub SetSideOfConflictSlicerCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S5FuncSetSideOfConflictSlicer)
End Sub

Private Sub Worksheet_Activate()
    Call ThisWorkbook.ExecuteProcedure(S5FuncWorksheetOpen)
End Sub

