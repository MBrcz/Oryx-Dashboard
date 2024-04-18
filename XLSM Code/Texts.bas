Attribute VB_Name = "Texts"
Option Explicit
' --------------------------------------------
' Project: Attack on Europe 2022
' Made by: Matthew Borcz
' GH Page: https://github.com/MBrcz
' --------------------------------------------
' - This page should consist of functions that represents the lithany of strings (flat text) -

' --------------------------------------------
' ---------------- BACKEND -------------------
' --------------------------------------------

' Date box text.
Public Function backend_dateInput(ByVal strDates As String) As String
    backend_dateInput = backend_dateInput1 & strDates & backend_dateInput2
End Function

Private Function backend_dateInput1() As String
    backend_dateInput1 = "Choose a date which you want to check data about." & vbCrLf & _
                          "Type an appriopriate number, where: " & "" & vbCrLf & _
                          "0: Exit Box" & Chr(10)
End Function

Private Function backend_dateInput2() As String
    backend_dateInput2 = "Current chosen date is: "
End Function

Public Function backend_dateError() As String
    backend_dateError = "Wrong input! Make sure that you are passing any INTEGER that is written in the question."
End Function
' END Date box text

' -------------------------------------------
' -------------- MILITARY FAMILY ------------
' -------------------------------------------

Public Function mili_EqChangingError() As String
    mili_EqChangingError = "Cannot find such equipment type! Select other option!"
End Function

Public Function wrong_ArgumentError() As String
    wrong_ArgumentError = "There has been supplied invalid argument to a function."
End Function

Public Function items_WrongQuery() As String
    items_WrongQuery = "Cannot find any valid item under this query in database." & vbCrLf & _
                       "Please, choose other options."
End Function

Public Function backend_RefreshInputBox() As String
    backend_RefreshInputBox = "Do you really want to refresh the data source?"
End Function

