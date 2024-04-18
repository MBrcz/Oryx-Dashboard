Attribute VB_Name = "Functions"
Option Explicit
' --------------------------------------------
' Project: Attack on Europe 2022
' Made by: Matthew Borcz
' GH Page: https://github.com/MBrcz
' --------------------------------------------

' --------------------------------------------------------------------------
' -------------------------------- UTILS -----------------------------------
' --------------------------------------------------------------------------

' It is not in the module with enums, becouse this enum is bound to one utility function,
' not logic of this project as a whole.
Public Enum eConverterMove
    eLeft = 1
    eRight = 2
    eTop = 3
    eBottom = 4
End Enum

' - SLICER MANAGMENT -
Public Sub SetCurrentSlicer(ByVal enumSlicerCall As SlicerCall, strValue As String)
    ' *** Sets the value of the current slicer in the project, depending on asked slicer.***
    
    ' Accepts:
    '   - enumSlicerCall [enum] - the number of the slicer that shall be changed.
    '   - strValue [string] - the value that is going to be passed into the slicer.
    ' Returns:
    '   - None
    
    Call SetPivotSlicer( _
              strSlicerName:=GetSlicerName(enumSlicerCall), _
              strTableName:=GetSlicerTable(enumSlicerCall), _
              strColumnName:=GetSlicerColumn(enumSlicerCall), _
              strValue:=strValue _
    )
    
    Call SetSlicerCurrentValue(enumSlicerCall, strValue)
    
End Sub

Public Sub ClearCurrentSlicer(ByVal enumSlicerCall As SlicerCall)
    ' *** Clears the content of the passed slicer. ***
    
    ' Accepts:
    '   - enumSlicerCall [enum] - the number of slicer that will be cleansed
    ' Returns:
    '   - None
    
    Call ClearPivotSlicer(GetSlicerName(enumSlicerCall))
    Call SetSlicerCurrentValue(enumSlicerCall, "All")
    
End Sub

Private Sub SetPivotSlicer(ByVal strSlicerName As String, ByVal strTableName As String, ByVal strColumnName As String, ByVal strValue As String)
    ' *** Sets the slicer for the known parameters. ***
    
    ' Accepts:
    '   - strSlicerName [string] - the name of the slicer that shall be used.
    '   - strTableName [string] - the name of the table that will be used.
    '   - strColumnName [string] - the column name that will be set
    '   - strValue [string] - the value of the slicer set.
    
    ' Returns:
    '   - None
    
    strTableName = "[" & strTableName & "]"
    strColumnName = "[" & strColumnName & "]"
    strValue = "[" & strValue & "]"
    
    ' Use an array to pass the value
    ActiveWorkbook.SlicerCaches(strSlicerName).VisibleSlicerItemsList = Array(strTableName & "." & strColumnName & ".&" & strValue)
End Sub

Private Sub ClearPivotSlicer(ByVal strSlicerName As String)
    ' ***Clears the Pivot Slicer completely ***
    
    ' Accepts:
    '   - strSlicerName [string] - the name of the slicer that will be cleared.
    ' Returns:
    '   - None
    
    ActiveWorkbook.SlicerCaches(strSlicerName).ClearManualFilter

End Sub

' - END SLICER MANGMENT -

' - VALIDATION LISTS -
Public Sub InjectContentToValidationList(ByVal rngTarget As Range, ByVal strExpression As String)
    ' *** Places the array in a range as a validation list ***
    
    ' Accepts:
    '   - rngTarget [range] - the range where the validation list shall be created
    '   - strExpression [string] - the expression which will be placed as a validation list.
    '                              It might be text splitted by commas or just an formula
    
    ' Returns:
    '   - None
    
    With rngTarget.Validation
        .Delete
        .Add xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=strExpression
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ShowInput = True
    End With

End Sub

Public Sub RemoveValidationListFromRange(ByVal rngTarget As Range)
    ' *** Removes the inserted validation list from selected range. ***
    
    ' Accepts:
    '   - rngTarget [range] - the range where validation shall be cleared.
    ' Returns:
    '   - None
    
    With rngTarget.Validation
        .Delete
    End With
    
End Sub
' - END VALIDATION LISTS -

' - Arrays -
Public Function ConvertRangeToArray(ByVal rngStartingRange As Range, ByVal enumConverterMove As eConverterMove, _
                                    Optional ByVal IsAddStartingRange As Boolean = False) As Variant
    ' *** This function is responsible for getting all elements meet***
    
    ' Accepts:
    '   - rngStartingRange [rng] - the range from which function will start moving.
    '   - enumConverterMove [enum] - the way in which function will get elements
    '   - IsAddStartingRange [bool] - True means that value of a rngStartingRange will be added, False otherwise.
    
    ' Returns:
    '   - array - new array.
    
    Dim arrResult() As Variant
    Dim rngNextCell As Range
    Dim intStart As Integer
    Dim i As Integer
    
    Set rngNextCell = rngStartingRange

    i = 0
    If IsAddStartingRange Then
        ReDim Preserve arrResult(0 To i)
        arrResult(i) = rngNextCell.Value
        i = i + 1
    End If
        
    Do While True
        Select Case enumConverterMove
            Case eLeft ' Lone left throws magical error cuz yes.
                Set rngNextCell = rngNextCell.Offset(0, -1)
                
            Case eRight
                Set rngNextCell = rngNextCell.Offset(0, 1)
                
            Case eTop
                Set rngNextCell = rngNextCell.Offset(-1, 0)
                
            Case eBottom
                Set rngNextCell = rngNextCell.Offset(1, 0)
        End Select
        
        If rngNextCell.Value = "" Then
            Exit Do
        
        Else
            ReDim Preserve arrResult(0 To i)
            arrResult(i) = rngNextCell.Value
            i = i + 1
        End If
    Loop
            
    ConvertRangeToArray = arrResult
        
End Function

Public Function IsElementInArray(ByVal arrArray As Variant, ByVal varElement As Variant) As Boolean
    ' ***
    ' Checks whether an element can be found in array or not.
    ' In case of passing empty variant it returns False.
    ' ***
    
    ' Accepts:
    '   - arrArray [array] - the array that will be tested correctness.
    '   - varElement [variant] - the element that will be searched for.
    
    ' Returns:
    '   - bool - True if element is in array otherwise false.
    
    Dim i As Integer
    
    If IsArray(arrArray) Then
        For i = LBound(arrArray) To UBound(arrArray)
            If arrArray(i) = varElement Then
                IsElementInArray = True
                Exit Function
            End If
        Next i
    End If
    
    IsElementInArray = False
    
End Function

Public Function IsArrayEmpty(ByRef arrArray() As Variant) As Boolean
    ' *** This function checks if passed array is empty or not ***
    
    ' Accepts:
    '  - arrArray [array] - array that shall be tested
    ' Returns:
    '   - boolean - True if it is empty otherwise false.

    If IsArray(arrArray) Then
        On Error Resume Next
        IsArrayEmpty = (LBound(arrArray) > UBound(arrArray))
        If err.Number <> 0 Then IsArrayEmpty = True
        On Error GoTo 0
    Else
        IsArrayEmpty = True
    End If

End Function
' - END ARRAYS -


' -------------------------------------------------------------------------
' ------------------------------ IMAGES -----------------------------------
' -------------------------------------------------------------------------

Public Sub PlaceImage(ByVal rngImageRange As Range, ByVal strPath As String, _
                      ByVal newImageName As String)
    ' ***
    '   Places Images In The Current Worksheet in the chosen range and sets the size according to range in question.
    '   Works only for images that have set arbitrary path.
    ' ***
    
    ' Accepts:
    '   - rngImageRange [range] - the range where image shall be placed
    '   - strPath [string] - the name of image that will be loaded.
    '   - newImageName [string] - the name of the new image that will be placed.
    ' Returns:
    '   - None
    Dim shpImg As Shape
    
    On Error Resume Next
    Set shpImg = ActiveSheet.Shapes.AddPicture(strPath, msoFalse, msoCTrue, 1, 1, 1, 1)
    Call PlaceImageInTheRange(shpImg, rngImageRange, newImageName)
    
End Sub

Private Sub PlaceImageInTheRange(ByVal shpImage As Shape, ByVal rngTarget As Range, ByVal strNewName As String)
    ' *** Places the image on the chosen range ***
    
    ' Accepts:
    '   - shpImage [Shape] - the image that shall be moved,
    '   - rngTarget [range] - the place where image will be copied to
    '   - strNewName [string] - the new name of the copied image.
    
    With shpImage
        .LockAspectRatio = msoFalse
        .Top = rngTarget.Top
        .Left = rngTarget.Left
        .Width = rngTarget.Width
        .Height = rngTarget.Height
        .Name = strNewName
    End With

End Sub

Public Sub RemoveImagesWithSignature(ByVal strSignature As String)
    ' *** Removes each shape that contains a name shown by strSignature ***
    
    ' Accepts:
    '   - strSignature [string] - the substring that shape contains in order to be removed.
    ' Returns:
    '   - None
    
    Dim shpImage As Shape
    
    For Each shpImage In ActiveSheet.Shapes
        If InStr(1, shpImage.Name, strSignature, vbTextCompare) Then
            shpImage.Delete
        End If
    Next shpImage
    
End Sub

Public Sub ClearTheLineInImages(ByVal strImageName As String)
    '*** Clears the line in the images, where shape object (image) contains specific string in it's name. ***
    
    ' Accepts:
    '   - strImageName [string] - the string of the letters that is in the image
    ' Returns:
    '   - None
    
    Dim shpImage As Shape
    
    For Each shpImage In ActiveSheet.Shapes
        If InStr(1, shpImage.Name, strImageName, vbTextCompare) Then
            shpImage.Line.Visible = msoFalse
        End If
    Next shpImage

End Sub

Public Sub SetTheLineInImages(ByVal strImageName As String)
    '*** Changes the visibility of the line in the images, where shape object (image) contains specific string in it's name. ***
    
    ' Accepts:
    '   - strImageName [string] - the string of the letters that is in the image
    ' Returns:
    '   - None
    
    Dim shpImage As Shape
    
    For Each shpImage In ActiveSheet.Shapes
        If InStr(1, shpImage.Name, strImageName, vbTextCompare) Then
            shpImage.Line.Visible = msoCTrue
        End If
    Next shpImage

End Sub

Public Sub InvokeFunctionWithinAnImage(ByVal strImageName As String, ByVal strCommandName As String)
    '*** Sets to the object that contains text strImageName the passed functions. ***
    
    ' Accepts:
    '   - strImageName [string] - the string of the letters that is in the image
    '   - strCommandName [string] - the name of the command that shall be passed.
    ' Returns:
    '   - None

    Dim shpImage As Shape
    
    For Each shpImage In ActiveSheet.Shapes
        If InStr(1, shpImage.Name, strImageName, vbTextCompare) Then
            shpImage.OnAction = strCommandName
        End If
    Next shpImage

End Sub
' - END IMAGE -

' - OTHER -
Public Sub ExecuteProcedure(ByVal strFunctionName As String, ParamArray arrParams() As Variant)
    ' ***
    ' Executes the passed function with variable ammounts of the arguments.
    ' Before execution it saps the enableevents and screenupdating worksheet properties
    ' in order to increase the speed of code. In case of either failure or sucess it
    ' restarts the settings to True.
    ' ***
    
    ' Accepts:
    '   - strFunctionName [str] - the name of the called function
    '   - arrParams() [array] - the arguments that are assosciated with the chosen function, can pass nothing.
    
    ' Returns:
    '   - None
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Application.Run strFunctionName, arrParams()
        
    If err.Number <> 0 Then
         MsgBox "There has been raised some error durning execution of " & strFunctionName & "!" & vbCrLf & vbCrLf & _
                "Number Error: " & err.Number & vbCrLf & "Description: " & err.Description, vbCritical
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Public Sub ZoomToLastVisibleColumnAndRow(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' It sets the zoom of the page as high as possible regarding
    ' to user screen settings
    ' ***
        
    Dim visibleRange As Range
    
    ' Application.ScreenUpdating = False
    
    Set visibleRange = ActiveSheet.Cells.SpecialCells(xlCellTypeVisible)
    visibleRange.Select
    
    ActiveWindow.Zoom = True
    ActiveSheet.Cells(1, 1).Select
    
    ' Application.ScreenUpdating = True

End Sub

Public Sub PreventScrollBarFromYoloSwitching(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function makes certain that bug from instant switching scrollbar form position Current Val ->
    ' Max val or vice versa won't happen
    ' ***
    
    ' Accepts:
    '   - None
    ' Returns:
    '   - None
    
    Application.Wait (Now() + 0.000005)

End Sub

Public Sub RevealAllWorksheets(Optional ByVal varProxy As Variant)
    ' *** This function is created for unhiding all worksheets that are created in this workbook. ***
    
    Dim wsSheetObject As Worksheet
    
    For Each wsSheetObject In ThisWorkbook.Worksheets
        wsSheetObject.Visible = xlSheetVisible
    Next wsSheetObject
    
End Sub
' - END OHTER -

