VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
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
Private Type tImageFlagParams
    ' *** Convienience ad hoc tool. Very usefull. ***
    strAreaRange As String
    strUrlPath As String
    strDemandedCountryName As String
    strImageSignature As String
    strImageName As String
End Type

Private Function GetTextBoxSlicersCodes() As Variant
    ' *** Gets the codes of all the slicers that have their input boxes in page. ***
    
    GetTextBoxSlicersCodes = Array(ItemBoxBranch, ItemBoxType, ItemBoxFamily, ItemBoxFullName, _
                                   ItemBoxCountryFullName, ItemBoxFateName)
End Function

' --------------------------------------------------------
' ------------------- PORTRAIR && FLAGS ------------------
' --------------------------------------------------------

' - IMAGE FLAGS -
Private Sub PopulateImageFlags()
    ' ***
    ' This function is responsible for populating the image flags in the project according to the known country
    ' ***

    Dim rngStartingRange As Range
    Dim typFlagParams As tImageFlagParams
    Dim i As Integer

    On Error Resume Next
    Set rngStartingRange = Range(Getters_Setters.GetItemOtherValue(ItemSearcherFlagStartRange))
    
    i = 0
    Do While Not rngStartingRange.Offset(i, 0).EntireRow.Hidden
        typFlagParams = InitializeImageParam(rngStartingRange.Offset(i, 0), i)
                
        ' If typFlagParams has not been instanciated (kill all images)
        If typFlagParams.strUrlPath = "" Then
            Call Functions.RemoveImagesWithSignature(Getters_Setters.GetItemOtherValue(ItemSearcherFlagSingature))
            Exit Sub
        
        ' Else check if is necessary to add image - if so then add
        ElseIf IsImageNeedToBeReplaced(typFlagParams.strImageName) Then
            Call Functions.RemoveImagesWithSignature(typFlagParams.strImageSignature & i / 2)
            Call Functions.PlaceImage( _
                           rngImageRange:=Range(typFlagParams.strAreaRange), _
                           strPath:=typFlagParams.strUrlPath, _
                           newImageName:=typFlagParams.strImageName _
            )
            
            Call ModifyCountryImageSizes(typFlagParams.strImageName, Range(typFlagParams.strAreaRange))
        End If
        
        i = i + 2
    Loop

End Sub

Private Function InitializeImageParam(ByVal rngSingleCell As Range, ByVal i As Integer) As tImageFlagParams
    ' ***
    ' This function is responsible for initializing parameters used in the PopulateImageFlags function.
    ' It exists only to decrease the loadout of the responsibility from metioned function.
    ' ***
    
    ' Accepts:
    '   - rngSingleCell [rng] - the cell that is the base for getting the parameters
    '   - i [int] - the iteration of the loop.
    
    ' Returns:
    '   - Type - the collection of intresting parameters.
    
    Dim typImageFlagParams As tImageFlagParams
    Dim rngMergedCells As Range
    
    Set rngMergedCells = rngSingleCell.MergeArea
    
    With typImageFlagParams
        .strImageSignature = Getters_Setters.GetItemOtherValue(ItemSearcherFlagSingature)
        .strAreaRange = rngMergedCells.Address
        .strUrlPath = Split(rngMergedCells.Cells(1, 1).Value, "||")(1)
        .strDemandedCountryName = Split(rngMergedCells.Cells(1, 1).Value, "||")(0)
        .strImageName = typImageFlagParams.strImageSignature & i / 2 & typImageFlagParams.strDemandedCountryName
    End With
        
    InitializeImageParam = typImageFlagParams
        
End Function

Private Function IsImageNeedToBeReplaced(ByVal strImageName As String) As Boolean
    ' *** Checks if there is a need for changing the image in the current cell. ***

    ' Accepts:
    '   - strImageName [string] - the demanded image name that will be searched for.
    ' Returns:
    '   - Boolean [bool] - True if image there is no such image, else True


    Dim shpFlagImage As Shape

    On Error Resume Next
    Set shpFlagImage = ActiveSheet.Shapes(strImageName)
    On Error GoTo 0

    If Not shpFlagImage Is Nothing Then
        IsImageNeedToBeReplaced = False
    Else
        IsImageNeedToBeReplaced = True
    End If

End Function

Private Sub ModifyCountryImageSizes(ByVal strImageName As String, ByVal rngCurrentCell As Range)
    ' *** Changes the size of the image inside the cell.***

    ' Accepts:
    '   - strImageName [string] - the name of the image that will be modified
    '   - rngCurrentCell [range] - the range in which the image shall be stored.
    ' Returns:
    '   - None

    Dim shpImage As Shape

    Set shpImage = ActiveSheet.Shapes(strImageName)

    With shpImage
        .Width = rngCurrentCell.Width - 20
        .Height = rngCurrentCell.Height - 20
        .Left = rngCurrentCell.Left + 10
        .Top = rngCurrentCell.Top + 10
    End With

End Sub

' - PORTRAIR SHENENIGANS -

Private Sub HandlePortrairOperations()
    ' *** This procedure is responsible for managing the portrairs operations in the page ***
    
    Dim strPortrairSingature As String
    
    strPortrairSingature = Getters_Setters.GetItemOtherValue(ItemSearcherPortrairSignature)
   
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
    
    Set rngPortrairRange = Range(GetItemOtherValue(ItemSearcherPortrairRange))
    strPortrairPath = Getters_Setters.GetItemOtherValue(ItemSearcherPortrairPath)
    
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
    
    strPortrairTextBoxName = Getters_Setters.GetItemOtherValue(ItemSearcherPortrairTextBoxName)
    strTextToPass = Getters_Setters.GetItemOtherValue(ItemSearcherTextBoxContent)

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

Private Sub InjectBookmarksToTextBoxes(ByVal enumItemSearcherBookmark As ItemSearcherBookmark)
    ' *** Places the bookmarks parameters in the text boxes in current module. ***
    
    ' Accepts:
    '   - enumItemSearcherBookmark [enum] - the number of the bookmark parameter that shall be set.
    ' Returns:
    '   - None
        
    Dim arrTextBoxCodes() As Variant
    Dim varTextBoxCode As Variant
    Dim rngTextBox As Range

    arrTextBoxCodes = GetTextBoxSlicersCodes()

    For Each varTextBoxCode In arrTextBoxCodes
        Set rngTextBox = Range(Getters_Setters.GetItemSearcherRange(varTextBoxCode))
        rngTextBox.Value = Getters_Setters.GetItemSearcherBookmarkParam(varTextBoxCode, enumItemSearcherBookmark)
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
        Set rngTextBox = Range(Getters_Setters.GetItemSearcherRange(varTextBoxCode))
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
    
    strTableName = Getters_Setters.GetItemOtherValue(ItemSearcherCorrectnessTable)
    
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
    
    strSlicerApplied = Getters_Setters.GetSlicerCurrentValue(itemsSideOfConflict)
    strSlicerToApply = Getters_Setters.GetItemOtherValue(enumItemOther)
    
    Call ClearSideOfConflictSlicer
    If strSlicerApplied <> strSlicerToApply Then
        Call Functions.SetCurrentSlicer(itemsSideOfConflict, strSlicerToApply)
        Call Functions.SetTheLineInImages(Getters_Setters.GetItemOtherValue(ItemSideOfConflictSignature) & enumItemOther)
    End If
    
End Sub

Private Sub ClearSideOfConflictSlicer()
    ' *** This function is responsible for removing the side of conflict slicer in the current page. ***
     
    Call Functions.ClearTheLineInImages(Getters_Setters.GetItemOtherValue(ItemSideOfConflictSignature))
    Call Functions.ClearCurrentSlicer(itemsSideOfConflict)

End Sub
' - END SIDE OF CONFLICT -

' - OTHER -
Private Sub RefreshObjectsBinds(Optional ByVal varProxy As Variant = 0)
    ' *** This function is responsible for refreshing the binds / text box sources. ***
    
    Dim shpObject As Shape
    
    Dim arrObjectNames() As Variant
    Dim arrFunctionsNames() As Variant
    Dim i As Integer
    
    arrObjectNames = Array(ItemSearcherRestartFiltersSignature, ItemSearcherSearchButtonSignature, _
                            ItemSearcherBookmarkSignature, ItemSideOfConflictSignature)
    arrFunctionsNames = Array(S4CallRestartFitlers, S4CallSearchQuery, _
                             S4CallInvokeBookmarks, S4CallSetSideOfConflictSlicer)
    
    For i = LBound(arrObjectNames) To UBound(arrObjectNames)
        Call Functions.InvokeFunctionWithinAnImage( _
                       strImageName:=Getters_Setters.GetItemOtherValue(arrObjectNames(i)), _
                       strCommandName:=Getters_Setters.GetFunctionName(arrFunctionsNames(i)) _
        )
    Next i
    
    Call InjectElementsToChooseToTextBoxes
    Call ThisWorkbook.UpdateScrollBar(itemsOne, S4CallUpdateScrollBar)
    Call PopulateImageFlags
    
End Sub

' - END OTHER -

' --------------------------------------------------------------
' -------------------- FUNCTIONS -------------------------------
' --------------------------------------------------------------


Public Sub UpdateTheLocalScrollbarFunction(Optional ByVal varProxy As Variant = 0)
    ' *** This function is responsible for handling the update of the scrollbar in local page. ***
    
    Call ThisWorkbook.ExecuteProcedure(FFuncPreventScrollBarFromSwitching, False)
    Call PopulateImageFlags
End Sub

Public Sub SearchQueryFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for handling searching the new query in the project
    ' ***
    
    If AreThoseSlicersPossibleToApply Then
        Call ClearTextBoxSlicers
        Call SetTextBoxSlicers
        Call HandlePortrairOperations
        Call ThisWorkbook.UpdateScrollBar(itemsOne, S4CallUpdateScrollBar)
        Call PopulateImageFlags
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
    Call ThisWorkbook.UpdateScrollBar(itemsOne, S4CallUpdateScrollBar)
    Call PopulateImageFlags
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
    Call ThisWorkbook.UpdateScrollBar(itemsOne, S4CallUpdateScrollBar)
    Call PopulateImageFlags
End Sub


Public Sub SetSideOfConflictSlicerFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for applying the side of conflicts slicers in the page.
    ' ***
    
    Dim intNum As Integer
    
    intNum = CInt(Right(Application.Caller, 1))
    Call InjectSideOfConflictSlicer(intNum)
    Call ThisWorkbook.UpdateScrollBar(itemsOne, S4CallUpdateScrollBar)
    Call PopulateImageFlags
End Sub

Public Sub WorksheetOpenFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for rebinding the binds in current page.
    ' Call ONLY directly from worksheet activate method.
    ' ***
    
    If ActiveSheet.Name = Getters_Setters.GetWorksheetsName(ItemsSearcher) Then
        Call RefreshObjectsBinds
        Call ThisWorkbook.ExecuteProcedure(FFuncZoomToVisibleCells, False)
        Call ThisWorkbook.ExecuteProcedure(S1FuncUpdateControlPanelObjects)
    End If
End Sub

' ----------------------------------------------------------------
' ------------------------- CALLERS ------------------------------
' ----------------------------------------------------------------

Public Sub PlaceBookmarkCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S4FuncInvokeBookmarks)
End Sub

Public Sub SearchQueryCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S4FuncSearchQuery)
End Sub

Public Sub RestartFiltersCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S4FuncRestartFilters)
End Sub

Public Sub SetSideOfConflictSlicerCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S4FuncSetSideOfConflictSlicer)
End Sub

Public Sub UpdateTheLocalScrollbarCall(Optional ByVal varProxy As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S4FuncUpdateScrollBar)
End Sub

Private Sub Worksheet_Activate()
    ThisWorkbook.ExecuteProcedure (S4FuncWorksheetOpen)
End Sub

