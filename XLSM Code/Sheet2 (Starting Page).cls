VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
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

Private Sub UpdateLocalScrollBars()
    ' *** This function is responsible for updating the parameters of the scrollbars in this page. ***
    
    Dim arrScrollBars As Variant
    Dim varScrollBar As Variant
    
    arrScrollBars = Array(startTop, startMid, startBot)
    
    For Each varScrollBar In arrScrollBars
        Call ThisWorkbook.UpdateScrollBar(varScrollBar, FFuncPreventScrollBarFromSwitching)
    Next varScrollBar
    
End Sub

Public Sub WorksheetOpenFunction(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' This function is responsible for rebinding the binds in current page.
    ' Call ONLY directly from worksheet activate method.
    ' ***
    
    If ActiveSheet.Name = GetWorksheetsName(StartingPage) Then
        Call ThisWorkbook.ExecuteProcedure(FFuncZoomToVisibleCells, False)
        Call UpdateLocalScrollBars
        Call ThisWorkbook.ExecuteProcedure(S1FuncUpdateControlPanelObjects, False)
    End If
End Sub

Private Sub Worksheet_Activate()
    Debug.Print "lol"
    Call ThisWorkbook.ExecuteProcedure(SStartFuncWorksheetOpen)
End Sub
