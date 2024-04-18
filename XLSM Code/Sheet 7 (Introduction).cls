VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
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

Private Sub RefreshObjectsBinds()
    ' *** This method is responsible for applying the binds inside the objects in the current page ***
    
    Dim arrFormulasIds() As Variant
    Dim arrObjectIds() As Variant
    Dim i As Integer
    
    arrObjectIds = Array(IntroRefresh, IntroExit, IntroModes)
    arrFormulasIds = Array(WbCallRefreshSource, S7FuncExitApp, WbfuncChangeWorkbookMode)
    
    For i = LBound(arrObjectIds) To UBound(arrObjectIds)
        Call Functions.InvokeFunctionWithinAnImage( _
                       strImageName:=Getters_Setters.GetOtherParam(arrObjectIds(i)), _
                       strCommandName:=Getters_Setters.GetFunctionName(arrFormulasIds(i)) _
        )
    Next i
    
End Sub

Public Sub ExitApplicationFunction(Optional ByVal varProxy = 0)
    ' *** As name implies, this procedure is handling the exitting of the programme. ***
        
    Application.DisplayAlerts = False
    Application.Quit
    Application.DisplayAlerts = True
End Sub

Public Sub WorksheetOpenFunction(Optional ByVal varProxy As Variant = 0)
    ' *** This method is responsible for binding an functions to objects. ***
        
    If ActiveSheet.Name = Getters_Setters.GetWorksheetsName(Introduction) Then
        Call RefreshObjectsBinds
        Call ThisWorkbook.ExecuteProcedure(FFuncZoomToVisibleCells, False)
        Call ThisWorkbook.ExecuteProcedure(S1FuncUpdateControlPanelObjects, False)
    End If
End Sub

Private Sub Worksheet_Activate()
    Call ThisWorkbook.ExecuteProcedure(S6FuncWorksheetOpen)
End Sub

