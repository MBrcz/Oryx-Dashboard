VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
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
' This page is responsible for managing the slicers and other operations
' in Military Overwiev Page of the project.

' -----------------------------------------------------------
' -------------------------- SLICER -------------------------
' -----------------------------------------------------------

Private Sub SetOverwievSlicer(ByVal intSlicerNum As Integer)
    ' *** Sets the value of the overwiev slicer in project ***
    
    ' Accepts:
    '   - intSlicerNum [integer] - the number of slicer that shall be set
    ' Returns:
    '   - None
    
    Dim strSlicerToApply As String
    Dim strSlicerCurrentValue As String
    
    strSlicerToApply = Getters_Setters.GetMilitaryOverwievSlicerValue(intSlicerNum)
    strSlicerCurrentValue = Getters_Setters.GetSlicerCurrentValue(milBranch)
    
    Call Functions.ClearTheLineInImages(Getters_Setters.GetSlicerSignature(milBranch))
    Call Functions.ClearCurrentSlicer(milBranch)
    
    If strSlicerToApply <> strSlicerCurrentValue Then
        Call Functions.SetCurrentSlicer(milBranch, strSlicerToApply)
        Call Functions.SetTheLineInImages(Getters_Setters.GetSlicerSignature(milBranch) & CStr(intSlicerNum))
    End If
    
End Sub

Private Sub UpdateShapesFunction()
    ' *** This function makes 100 % sure, that the slicers are bound to an function. ***
    
    Call Functions.InvokeFunctionWithinAnImage(strImageName:=Getters_Setters.GetSlicerSignature(milBranch), _
                                               strCommandName:=Getters_Setters.GetFunctionName(S2CallSetOverwievSlicer) _
    )
End Sub


' -----------------------------------------------
' --------------- FUNCTIONS TO CALL -------------
' -----------------------------------------------
' Those function shall be wrapped - do not call them directly by user.
' The parameter varProxy makes sure that no one shall bind them by accident
' to the GUI.

Public Sub SetOverviewSlicerFunction(Optional ByVal varProxy As Variant = 0)
    ' *** This function is responsible for setting and changing the slicers in current page. ***
    Dim intNum As Integer
    
    intNum = CInt(Right(Application.Caller, 1))
    Call SetOverwievSlicer(intNum)
End Sub

Public Sub WorksheetOpenFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for rebinding the binds in current page.
    ' Call ONLY directly from worksheet activate method.
    ' ***
    
    If ActiveSheet.Name = Getters_Setters.GetWorksheetsName(MilitaryOverwiev) Then
        Call ThisWorkbook.ExecuteProcedure(FFuncZoomToVisibleCells, False)
        Call UpdateShapesFunction
        Call ThisWorkbook.ExecuteProcedure(S1FuncUpdateControlPanelObjects)
    End If
End Sub

' -----------------------------------------------
' ------------------ CALLBACKS-------------------
' -----------------------------------------------

Public Sub SetOverviewSlicerCall(Optional ByVal varProxy As Variant = 0)
    ' See more: Functions.ExecuteProcedure or ThisWorkbook.
    Call ThisWorkbook.ExecuteProcedure(S2FuncSetOverwievSlicer)
End Sub

Private Sub Worksheet_Activate()
    Call ThisWorkbook.ExecuteProcedure(S2FuncWorksheetOpen)
End Sub
