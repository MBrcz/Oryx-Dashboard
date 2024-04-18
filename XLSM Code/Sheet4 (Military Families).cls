VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
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

Private Function GetEquipmentTypeStartingRange() As Range
    ' This function is setted up here, becouse it is unique.
    ' For some magical reason, without Application at the start, object range will an throw an error.
    ' Why? I do not know, oh great microsoft...
    ' Indeed, mystical are the ways sometimes VBA works.
    Set GetEquipmentTypeStartingRange = Application.Range(Getters_Setters.GetOtherParam(FamilyEquipmentTypeStartRange))
End Function

' --------------------------------------------------------------
' --------------------------- SLICERS --------------------------
' --------------------------------------------------------------

Private Function IsEquipmentTypeSlicerCorrect(ByVal strSlicerToApply As String) As Boolean
    ' *** This function is responsible for making sure that the passed by user equipment type slicer exists ***
    
    ' Accepts:
    '   - strSlicerToApply [string] - the slicer that should be applied.
    ' Returns:
    '   - bool - True means it is correct otherwise it is incorrect
    
    Dim varEquipmentTypeToChoose As Variant

    varEquipmentTypeToChoose = Functions.ConvertRangeToArray( _
                                          rngStartingRange:=GetEquipmentTypeStartingRange(), _
                                          enumConverterMove:=eBottom, _
                                          IsAddStartingRange:=True _
    )

    If Functions.IsElementInArray(varEquipmentTypeToChoose, strSlicerToApply) Then
        IsEquipmentTypeSlicerCorrect = True
    Else
        IsEquipmentTypeSlicerCorrect = False
    End If
    
End Function
' - END EQUIPMENT SLICER -

' - FAMILY SLICER -
Public Sub SetFamilySlicer(ByVal intSlicerNum As Integer)
    ' *** Sets the content of the family slicer by chosen number ***
    
    ' Accepts:
    '   - intSlicerNum [integer] - the number of the slicer that shall be set
    ' Returns:
    '   - None
    
    Dim strCurrentFamilySlicer As String
    Dim strFamilySlicerToApply As String
    
    strCurrentFamilySlicer = Getters_Setters.GetSlicerCurrentValue(familyFamilyName)
    strFamilySlicerToApply = Getters_Setters.GetFamilyObjectDescription(intSlicerNum)
    
    Call Functions.ClearTheLineInImages(Getters_Setters.GetSlicerSignature(familyFamilyName))
    Call Functions.ClearCurrentSlicer(familyFamilyName)
    
    If strFamilySlicerToApply <> strCurrentFamilySlicer Then
        Call Functions.SetCurrentSlicer(familyFamilyName, strFamilySlicerToApply)
        Call Functions.SetTheLineInImages(Getters_Setters.GetSlicerSignature(familyFamilyName) & CStr(intSlicerNum))
    End If
    
End Sub

' - END FAMILY SLICER -

' ---------------------------------------------------------------------------
' --------------------------- GRAPHICAL OBJECTS -----------------------------
' ---------------------------------------------------------------------------

Private Sub UpdatePageVisualContent()
    ' *** This method is responsible for udpating the visual side of this page. ***
    
    Dim arrFamilyObjects As Variant
    Dim varFamilyObject As Variant
    
    arrFamilyObjects = Array(FamilyObject.First, FamilyObject.Second, FamilyObject.Third, FamilyObject.Fourth)
    
    Call Functions.RemoveImagesWithSignature(Getters_Setters.GetSlicerSignature(familyFamilyName))
    Call Functions.RemoveImagesWithSignature(Getters_Setters.GetOtherParam(FamilySlicerTextBoxName))
    
    For Each varFamilyObject In arrFamilyObjects
        varFamilyObject = CInt(varFamilyObject) ' Do not do this at home. Dangerous code ahead.
        
        If IsFamilyObjectNeedToBeAdded(varFamilyObject) Then
            Call InsertNewImage(varFamilyObject)
            Call InsertNewTextBox(varFamilyObject)
        End If
    Next varFamilyObject
    
End Sub

Private Function IsFamilyObjectNeedToBeAdded(ByVal intSlicerNum As Integer) As Boolean
    ' *** This method checks if there is a necessity for adding the another visual object in the page. **
    
    ' Accepts:
    '   - intSlicerNum [integer] - the number of the slicer that will be chcecked.
    ' Returns:
    '   - None
    
    If Getters_Setters.GetFamilyObjectDescription(intSlicerNum) <> "" Then
        IsFamilyObjectNeedToBeAdded = True
    Else
        IsFamilyObjectNeedToBeAdded = False
    End If
    
End Function

' - IMAGE -
Private Sub InsertNewImage(ByVal intSlicerNum As Integer)
    ' *** Inserts the new image object to the page. ***
    
    ' Accepts:
    '   = intSlicerNum [integer] - the number of family object that the slicer num refers to.
    ' Returns:
    '   - None
    
    Dim strImageName As String
    
    strImageName = Getters_Setters.GetSlicerSignature(familyFamilyName) & CStr(intSlicerNum)
    Call Functions.PlaceImage( _
                   rngImageRange:=Range(Getters_Setters.GetFamilyImagePosition(intSlicerNum)), _
                   strPath:=Getters_Setters.GetFamilyImageUrl(intSlicerNum), _
                   newImageName:=strImageName _
    )
    
    Call ModifyImageProperties(ActiveSheet.Shapes(strImageName))
    
End Sub

Private Sub ModifyImageProperties(ByVal shpImage As Shape)
    ' *** This procedure is responsible for modyfing the newly inserted image properties ***
    
    ' Accepts:
    '   - shpImage [shape] - the image object that shall be modified.
    ' Returns:
    '   - None
    
    With shpImage
        ' Change apperance
        With shpImage.Line
            .Weight = 3
            .ForeColor.RGB = RGB(255, 0, 0)
            .Visible = False
        End With
        
        .OnAction = Getters_Setters.GetFunctionName(S3CallSetFamilySlicer)
    End With

End Sub
' - END IMAGE -

Private Sub InsertNewTextBox(ByVal intSlicerNum As Integer)
    ' *** Inserts the textbox object and modifies it's parameters. ***
    
    ' Accepts:
    '   = intSlicerNum [integer] - the number of the family object that the slicer num refers to.
    ' Returns:
    '   = None
    
    Dim rngTextBoxRange As Range
    Dim shpTextBox As Shape
    
    Set rngTextBoxRange = Range(Getters_Setters.GetFamilyTextBoxRange(intSlicerNum))
    Set shpTextBox = Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 20, 10)
    
    With shpTextBox
        ' Size
        .LockAspectRatio = msoFalse
        .Top = rngTextBoxRange.Top
        .Left = rngTextBoxRange.Left
        .Height = rngTextBoxRange.Height
        .Width = rngTextBoxRange.Width
        
        With .TextFrame
            With .Characters
                .Text = Getters_Setters.GetFamilyObjectDescription(intSlicerNum)
                With .Font
                    .Bold = msoTrue
                    .Size = 10
                    .Name = "Times New Roman"
                End With
            End With
            
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlHAlignCenter
        End With
        
        .Name = Getters_Setters.GetOtherParam(FamilySlicerTextBoxName) & CStr(intSlicerNum)
        .OnAction = Getters_Setters.GetFunctionName(S3CallSetFamilySlicer)
    End With

End Sub

' --------------------------------------------------------------
' --------------------- UPDATES PAGE ---------------------------
' --------------------------------------------------------------

Private Sub UpdateThePage()
    ' *** This procedure is responsible for refreshing the page slicers, every time the page is opened. ***
    
    ' Generates crash for some magical reason.
    Call InsertTheValidationBox

    Call Functions.InvokeFunctionWithinAnImage( _
                   strImageName:=Getters_Setters.GetOtherParam(FamilyEquipmentTypeSlicer), _
                   strCommandName:=Getters_Setters.GetFunctionName(S3CallSetEquipmentTypeSlicer) _
    )
    
    Call Functions.InvokeFunctionWithinAnImage( _
                   strImageName:=Getters_Setters.GetSlicerSignature(familyFamilyName), _
                   strCommandName:=Getters_Setters.GetFunctionName(S3CallSetFamilySlicer) _
    )
    

End Sub

' Generates error, function invalid.
Private Sub InsertTheValidationBox()
    ' ***
    ' This function is responsible for adding the content to the validation boxes in
    ' range that contains the equipment slicer potencial input.
    ' ***

    Dim rngValidationRange As Range
    Dim arrEquipmentTypeChoices As Variant
    Dim strValidationList As String

    Dim i As Integer

    Set rngValidationRange = Range(Getters_Setters.GetOtherParam(FamilyEquipmentTypeCellRange))
    arrEquipmentTypeChoices = Functions.ConvertRangeToArray( _
                                        rngStartingRange:=GetEquipmentTypeStartingRange(), _
                                        enumConverterMove:=eBottom, _
                                        IsAddStartingRange:=True _
    )
    
    ' Get rid of the problem with commas.
    
    i = 0
    For i = LBound(arrEquipmentTypeChoices) To UBound(arrEquipmentTypeChoices)
        If InStr(1, arrEquipmentTypeChoices(i), ",", vbTextCompare) Then
            arrEquipmentTypeChoices(i) = Replace(arrEquipmentTypeChoices(i), ",", ";", , , vbTextCompare)
        End If
    Next i
    
    Debug.Print rngValidationRange.Address
    Call Functions.InjectContentToValidationList( _
                   rngTarget:=rngValidationRange, _
                   strExpression:=Application.WorksheetFunction.TextJoin(",", True, arrEquipmentTypeChoices) _
    )

End Sub

' --------------------------------------------------------------
' --------------------- FUNCTIONS TO CALL ----------------------
' --------------------------------------------------------------
' Do not call directly those function by name!
' They will be later wrapped and called
' varParam is used as an mechanism against binding images to the objects from GUI.

Public Sub SwitchEquipmentTypeSlicerFunction(Optional ByVal varParam As Variant = 0)
    ' *** This procedure is responsible for changing the equipment type slicer in the project.***
    
    Dim strSlicerToApply As String
    
    strSlicerToApply = Getters_Setters.GetOtherParam(FamilyEqTypeToPass)
    
    If IsEquipmentTypeSlicerCorrect(strSlicerToApply) Then
        Call Functions.ClearCurrentSlicer(familyFamilyName)
        Call Functions.SetCurrentSlicer(familyEquipmentType, strSlicerToApply)
        Call UpdatePageVisualContent
    Else
        MsgBox Texts.mili_EqChangingError(), vbCritical + vbOKOnly, "Err"
        Call InsertTheValidationBox
    End If
    
End Sub

Public Sub SwitchFamilySlicerFunction(Optional ByVal varParam As Variant = 0)
    ' *** Handles changing the family slicer and it's machinations. ***
    
    Dim intFamilyNum As Integer
    
    intFamilyNum = CInt(Right(Application.Caller, 1))
    Call SetFamilySlicer(intFamilyNum)
    
End Sub

Public Sub HideEquipmentTypeValidationBoxFunction(Optional ByVal varParam As Variant = 0)
    ' ***
    ' This is quick hotfix for the fact, that validation box tends to generate error due to
    ' too much text inserted into them.
    ' Call this from workbook_close event.
    ' ***
    
    Call Functions.InjectContentToValidationList( _
                   rngTarget:=Range(Getters_Setters.GetOtherParam(FamilyEquipmentTypeCellRange)), _
                   strExpression:="A,B" _
    )
                   
    
End Sub

Public Sub WorksheetOpenFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for rebinding the binds in current page.
    ' Call ONLY directly from worksheet activate method.
    ' ***
    
    If ActiveSheet.Name = Getters_Setters.GetWorksheetsName(MilitaryFamilies) Then
        Call ThisWorkbook.ExecuteProcedure(FFuncZoomToVisibleCells, False)
        Call UpdateThePage
        Call ThisWorkbook.ExecuteProcedure(S1FuncUpdateControlPanelObjects)
    End If
End Sub

' --------------------------------------------------------------
' ---------------------- CALL ----------------------------------
' --------------------------------------------------------------
' Use those functions as callbacks in the objects in this sheet.

Public Sub SwitchFamilySlicerCall(Optional ByVal varParam As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S3FuncSetFamilySlicer)
End Sub

Public Sub SwitchEquipmentTypeSlicerCall(Optional ByVal varParam As Variant = 0)
    Call ThisWorkbook.ExecuteProcedure(S3FuncSetEquipmentTypeSlicer)
End Sub

Private Sub Worksheet_Activate()
    Call ThisWorkbook.ExecuteProcedure(S3FuncWorksheetOpen)
End Sub
