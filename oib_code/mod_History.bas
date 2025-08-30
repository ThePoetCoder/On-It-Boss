Attribute VB_Name = "mod_History"
Option Explicit
Option Private Module

Function History_GetCodeById( _
    ByVal wb As Workbook, _
    ByVal id As Long, _
    ByRef outTitle As String, _
    ByRef outLang As String _
) As String
    ' Returns the stored M code by history ID, also providing the saved title and language.
    Dim ws As Worksheet: Set ws = EnsureHistorySheet(wb)
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, mod_Config.hcID).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If CLng(ws.Cells(r, mod_Config.hcID).Value) = id Then
            outTitle = CStr(ws.Cells(r, mod_Config.hcTitle).Value)
            outLang = CStr(ws.Cells(r, mod_Config.hcLanguage).Value)
            History_GetCodeById = CStr(ws.Cells(r, mod_Config.hcCode).Value)
            Exit Function
        End If
    Next r
End Function

Sub History_DeleteById( _
    ByVal wb As Workbook, _
    ByVal id As Long _
)
    ' Deletes a single history row by its ID from the _OIBHistory sheet.
    Dim ws As Worksheet: Set ws = EnsureHistorySheet(wb)
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, hcID).End(xlUp).Row
    Dim r As Long
    Application.ScreenUpdating = False
    For r = lastRow To 2 Step -1
        If CLng(ws.Cells(r, hcID).Value) = id Then
            ws.rows(r).Delete
            Exit For
        End If
    Next r
    Application.ScreenUpdating = True
End Sub

Sub Panel_FillListForTable( _
    ByVal frm As Object, _
    ByVal lo As ListObject _
)
    ' Populates the uf_Panel.lst_Prev listbox with history titles (hiding the ID in column 0).
    On Error GoTo SafeExit
    frm.lst_Prev.clear
    If lo Is Nothing Then Exit Sub

    Dim arr As Variant
    arr = mod_History.History_ListForTable(lo.Parent.Parent, lo.name)
    If IsEmpty(arr) Then Exit Sub

    ' Show Title; keep ID in a hidden column for lookup
    frm.lst_Prev.ColumnCount = 2
    ' Show Title only (second column), hide first (ID)
    frm.lst_Prev.ColumnWidths = "0 pt;250 pt"
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        frm.lst_Prev.AddItem
        frm.lst_Prev.List(frm.lst_Prev.ListCount - 1, 0) = arr(i, 1) ' ID (hidden)
        frm.lst_Prev.List(frm.lst_Prev.ListCount - 1, 1) = arr(i, 2) ' Title (visible)
    Next i
SafeExit:
End Sub

Function EnsureHistorySheet( _
    ByVal wb As Workbook _
) As Worksheet
    ' Creates the History sheet if missing and hides it; writes column headers if needed.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(mod_Config.HISTORY_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = mod_Config.HISTORY_SHEET_NAME
        ws.Visible = xlSheetHidden
        ws.Range("A1:G1").Value = Array("ID", "TableName", "QueryName", "Title", "Language", "Code", "CreatedAt")
        ws.Cells.WrapText = False
    End If
    Set EnsureHistorySheet = ws
End Function

Public Sub History_EnsureSeedForTable( _
    ByVal wb As Workbook, _
    ByVal lo As ListObject _
)
    Dim ws As Worksheet: Set ws = EnsureHistorySheet(wb)
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, hcID).End(xlUp).Row

    Dim hasAny As Boolean, r As Long
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, hcTableName).Value), lo.name, vbTextCompare) = 0 Then
            hasAny = True
            Exit For
        End If
    Next

    If hasAny Then Exit Sub

    Dim queryName As String: queryName = mod_String.SanitizeName(lo.name)
    Dim seedM As String
    Dim existingM As String
    existingM = mod_SchemaQuery.GetExistingQueryFormula(wb, queryName)
    
    If Len(Trim$(existingM)) > 0 Then
        seedM = existingM                   ' << use actual formula (works for ANY source)
    Else
        seedM = mod_Core.BuildSeedM(lo.name) ' << fallback to Excel.CurrentWorkbook stub
    End If
    
    Call History_Add(wb, lo.name, queryName, "Load '" & lo.name & "'", "m", seedM)
End Sub

Function NextHistoryId( _
    ByVal ws As Worksheet _
) As Long
    ' Returns the next monotonically increasing history ID.
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, hcID).End(xlUp).Row
    If lastRow < 2 Then
        NextHistoryId = 1
    Else
        NextHistoryId = CLng(ws.Cells(lastRow, hcID).Value) + 1
    End If
End Function

Function History_Add( _
    ByVal wb As Workbook, _
    ByVal tableName As String, _
    ByVal queryName As String, _
    ByVal title As String, _
    ByVal language As String, _
    ByVal mcode As String _
) As Long
    ' Appends a history row with metadata and M code; returns the inserted row index.
    Dim ws As Worksheet, r As Long
    Set ws = EnsureHistorySheet(wb)
    r = ws.Cells(ws.rows.Count, hcID).End(xlUp).Row + 1
    If r = 2 And Len(ws.Cells(1, 1).Value) = 0 Then r = 2 ' ensure headers exist

    Dim newId As Long
    newId = NextHistoryId(ws)

    ws.Cells(r, mod_Config.hcID).Value = newId
    ws.Cells(r, mod_Config.hcTableName).Value = tableName
    ws.Cells(r, mod_Config.hcQueryName).Value = queryName
    ws.Cells(r, mod_Config.hcTitle).Value = title
    ws.Cells(r, mod_Config.hcLanguage).Value = language
    ws.Cells(r, mod_Config.hcCode).Value = mcode
    ws.Cells(r, mod_Config.hcCreatedAt).Value = Now
    ws.Range(ws.Cells(r, mod_Config.hcID), ws.Cells(r, mod_Config.hcCreatedAt)).WrapText = False

    History_Add = r
End Function

Function History_ListForTable( _
    ByVal wb As Workbook, _
    ByVal tableName As String _
) As Variant
    ' Returns a [n x 3] array of {ID, Title, CreatedAt} for a given table, oldest?newest.
    Dim ws As Worksheet: Set ws = EnsureHistorySheet(wb)
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, hcID).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim tmp As Collection: Set tmp = New Collection
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, mod_Config.hcTableName).Value), tableName, vbTextCompare) = 0 Then
            ' collect as "oldest to newest" (natural order since we always append)
            tmp.Add Array(ws.Cells(r, mod_Config.hcID).Value, ws.Cells(r, mod_Config.hcTitle).Value, ws.Cells(r, mod_Config.hcCreatedAt).Value)
        End If
    Next r

    If tmp.Count = 0 Then Exit Function

    Dim outArr() As Variant
    ReDim outArr(1 To tmp.Count, 1 To 3)
    Dim i As Long
    For i = 1 To tmp.Count
        outArr(i, 1) = tmp(i)(0) ' ID
        outArr(i, 2) = tmp(i)(1) ' Title
        outArr(i, 3) = tmp(i)(2) ' CreatedAt
    Next i
    History_ListForTable = outArr
End Function

