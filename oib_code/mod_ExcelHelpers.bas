Attribute VB_Name = "mod_ExcelHelpers"
Option Explicit
Option Private Module

Function ActiveListObjectOrNothing() As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    On Error GoTo 0

    If lo Is Nothing Then
        If ActiveSheet.ListObjects.Count >= 1 Then
            Set lo = ActiveSheet.ListObjects(1)
        Else
            Set ActiveListObjectOrNothing = Nothing
            Exit Function
        End If
    End If
    Set ActiveListObjectOrNothing = lo
End Function

