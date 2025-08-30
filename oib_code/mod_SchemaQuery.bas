Attribute VB_Name = "mod_SchemaQuery"
Option Explicit
Option Private Module

Sub ApplyMToTable( _
    ByVal lo As ListObject, _
    ByVal mcode As String, _
    Optional ByVal title As String = "" _
)
    If lo Is Nothing Then
        MsgBox "No table selected.", vbExclamation
        Exit Sub
    End If

    Dim wb As Workbook: Set wb = lo.Parent.Parent
    Dim qName As String: qName = mod_String.SanitizeName(lo.name)

    Dim hadQuery As Boolean: hadQuery = QueryExists(wb, qName)
    Dim oldFormula As String
    If hadQuery Then oldFormula = wb.Queries(qName).Formula

    On Error GoTo Rollback

    If hadQuery Then
        wb.Queries(qName).Formula = mcode
    Else
        wb.Queries.Add name:=qName, Formula:=mcode
    End If
    'ensure it’s always bound to a sheet / displayed
    EnsureQueryLoadedToSheet wb, qName  ' create/bind if needed

    ' Force a synchronous refresh of all dependents; will error if M is invalid
    RefreshAllBoundTables wb, qName

    If Len(title) > 0 Then
        ' optional: Application.StatusBar = "Applied: " & title
    End If
    On Error GoTo 0
    Exit Sub

Rollback:
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    If hadQuery Then
        wb.Queries(qName).Formula = oldFormula
        RefreshAllBoundTables wb, qName
    Else
        ' we created a brand-new query that failed; remove it and any sheet we made
        wb.Queries(qName).Delete
    End If
    On Error GoTo 0
    MsgBox "Your M code failed to apply/refresh and was rolled back." & vbCrLf & vbCrLf & _
           "Error: " & errMsg, vbExclamation
End Sub

Public Function TryExtractOriginTableName( _
    ByVal mcode As String _
) As String
    ' Find: Excel.CurrentWorkbook(){[Name="TableName"]}[Content]
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = "Excel\.CurrentWorkbook\(\)\{\[Name=""([^""]+)""\]\}\[Content\]"
    Dim matches As Object
    Set matches = re.Execute(mcode)
    If matches.Count > 0 Then
        TryExtractOriginTableName = matches(0).SubMatches(0)
    End If
End Function

Public Function FindListObjectByName( _
    ByVal wb As Workbook, _
    ByVal tableName As String _
) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                Set FindListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Public Function BuildOriginSchemaJson( _
    ByVal wb As Workbook, _
    ByVal mcode As String _
) As String
    ' Best-effort: parse origin table from M and emit its schema JSON.
    Dim tName As String: tName = TryExtractOriginTableName(mcode)
    If Len(tName) = 0 Then
        BuildOriginSchemaJson = "{}"
        Exit Function
    End If

    Dim lo As ListObject
    Set lo = FindListObjectByName(wb, tName)
    If lo Is Nothing Then
        BuildOriginSchemaJson = "{" & mod_String.QUOTE_CHAR & "tableName" & mod_String.QUOTE_CHAR & ":" & _
            mod_String.QUOTE_CHAR & mod_String.EscapeJson(tName) & mod_String.QUOTE_CHAR & "," & _
            mod_String.QUOTE_CHAR & "columns" & mod_String.QUOTE_CHAR & ":[]," & _
            mod_String.QUOTE_CHAR & "sample" & mod_String.QUOTE_CHAR & ":[]}"
    Else
        BuildOriginSchemaJson = BuildTableSchemaJson(lo)
    End If
End Function

Function BuildTableSchemaJson( _
    ByVal tbl As ListObject _
) As String
    ' Builds a compact schema/sample JSON from a ListObject: {tableName, columns[], sample[][]}.
    Dim json As String
    Dim colIndex As Long
    Dim maxRows As Long
    Dim rowIndex As Long

    json = "{" & mod_String.QUOTE_CHAR & "tableName" & mod_String.QUOTE_CHAR & ":" & mod_String.QUOTE_CHAR & mod_String.EscapeJson(tbl.name) & mod_String.QUOTE_CHAR & "," & mod_String.QUOTE_CHAR & "columns" & mod_String.QUOTE_CHAR & ":["
    For colIndex = 1 To tbl.ListColumns.Count
        If colIndex > 1 Then json = json & ","
        json = json & "{" & mod_String.QUOTE_CHAR & "name" & mod_String.QUOTE_CHAR & ":" & mod_String.QUOTE_CHAR & mod_String.EscapeJson(tbl.ListColumns(colIndex).name) & mod_String.QUOTE_CHAR & "}"
    Next colIndex
    json = json & "]," & mod_String.QUOTE_CHAR & "sample" & mod_String.QUOTE_CHAR & ":["

    If Not tbl.DataBodyRange Is Nothing Then
        maxRows = WorksheetFunction.Min(3, tbl.DataBodyRange.rows.Count)
    Else
        maxRows = 0
    End If

    If maxRows > 0 Then
        For rowIndex = 1 To maxRows
            Dim col2 As Long
            If rowIndex > 1 Then json = json & ","
            json = json & "["
            For col2 = 1 To tbl.ListColumns.Count
                If col2 > 1 Then json = json & ","
                json = json & mod_String.QUOTE_CHAR & mod_String.EscapeJson(CStr(tbl.DataBodyRange.Cells(rowIndex, col2).Value)) & mod_String.QUOTE_CHAR
            Next col2
            json = json & "]"
        Next rowIndex
    End If

    json = json & "]}"
    BuildTableSchemaJson = json
End Function

Function GetExistingQueryFormula( _
    ByVal wb As Workbook, _
    ByVal qName As String _
) As String
    ' Returns the existing WorkbookQuery formula by name (empty on missing).
    On Error Resume Next
    GetExistingQueryFormula = wb.Queries(qName).Formula
    On Error GoTo 0
End Function

Function QueryExists( _
    ByVal wb As Workbook, _
    ByVal qName As String _
) As Boolean
    ' True if a query with the given name exists in the workbook.
    Dim q As WorkbookQuery
    For Each q In wb.Queries
        If StrComp(q.name, qName, vbTextCompare) = 0 Then
            QueryExists = True
            Exit Function
        End If
    Next q
End Function

Sub LoadQueryToNewSheet( _
    ByVal wb As Workbook, _
    ByVal qName As String _
)
    ' Creates a new worksheet and binds a ListObject to the given query via Mashup OLE DB.
    Dim ws As Worksheet
    Dim oListObj As ListObject
    Dim mashupConn As String

    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.name = UniqueSheetName(wb, Left$(qName, 28))

    mashupConn = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName & ";Extended Properties="""""

    Set oListObj = ws.ListObjects.Add( _
        SourceType:=xlSrcExternal, _
        Source:=Array(mashupConn), _
        Destination:=ws.Range("A1") _
    )

    oListObj.name = mod_String.SanitizeTableName(qName)

    With oListObj.QueryTable
        .CommandType = xlCmdSql
        .CommandText = "SELECT * FROM [" & qName & "]"
        .AdjustColumnWidth = True
        .RefreshStyle = xlInsertDeleteCells ' important for schema changes
        .BackgroundQuery = False            ' synchronous
        .EnableRefresh = True
        .Refresh BackgroundQuery:=False     ' do an initial bind/refresh now
    End With
End Sub

Function IsListObjectBoundToQuery( _
    ByVal lo As ListObject, _
    ByVal qName As String _
) As Boolean
    ' Checks if a ListObject’s connection targets the given Mashup Location=QueryName.
    On Error Resume Next
    If lo Is Nothing Then Exit Function
    If lo.QueryTable Is Nothing Then Exit Function
    If lo.QueryTable.WorkbookConnection Is Nothing Then Exit Function
    If lo.QueryTable.WorkbookConnection.OLEDBConnection Is Nothing Then Exit Function

    Dim s As String
    s = lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection
    IsListObjectBoundToQuery = (InStr(1, s, "Microsoft.Mashup.OleDb.1", vbTextCompare) > 0) _
                            And (InStr(1, s, "Location=" & qName, vbTextCompare) > 0)
    On Error GoTo 0
End Function

Sub RefreshAllBoundTables( _
    ByVal wb As Workbook, _
    ByVal qName As String _
)
    ' Refreshes all ListObjects that are bound to the specified Mashup query.
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If IsListObjectBoundToQuery(lo, qName) Then
                With lo.QueryTable
                    On Error GoTo RefreshFail
                    .BackgroundQuery = False
                    .RefreshStyle = xlInsertDeleteCells
                    .Refresh BackgroundQuery:=False
                    On Error GoTo 0
                End With
                GoTo NextLo
                
RefreshFail:
                ' Bubble up so ApplyMToTable can rollback
                Err.Raise Err.Number, "RefreshAllBoundTables", Err.Description
                
NextLo:
            End If
        Next lo
    Next ws
End Sub

Sub EnsureQueryAndSheetForTable( _
    ByVal lo As ListObject _
)
    'ensures a query and sheet to display query both exist
    If lo Is Nothing Then Exit Sub
    Dim wb As Workbook: Set wb = lo.Parent.Parent
    Dim qName As String: qName = mod_String.SanitizeName(lo.name)

    If QueryExists(wb, qName) Then
        ' Query exists – make sure it’s actually bound to a sheet
        EnsureQueryLoadedToSheet wb, qName
        Exit Sub
    End If

    ' No query yet – create the seed M from the table and bind it
    Dim seedM As String
    seedM = mod_Core.BuildSeedM(lo.name)

    wb.Queries.Add name:=qName, Formula:=seedM
    EnsureQueryLoadedToSheet wb, qName
    RefreshAllBoundTables wb, qName
End Sub

Sub EnsureQueryLoadedToSheet( _
    ByVal wb As Workbook, _
    ByVal qName As String _
)
    ' Ensures at least one worksheet is bound to the given query, creating one if missing.
    Dim ws As Worksheet, lo As ListObject, found As Boolean

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If IsListObjectBoundToQuery(lo, qName) Then
                found = True
                Exit For
            End If
        Next lo
        If found Then Exit For
    Next ws

    If Not found Then
        LoadQueryToNewSheet wb, qName
    End If
End Sub

Public Function GetFirstBoundListObject( _
    ByVal wb As Workbook, _
    ByVal qName As String _
) As ListObject
    'find the first ListObject bound to a given query name
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If IsListObjectBoundToQuery(lo, qName) Then
                Set GetFirstBoundListObject = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Public Sub EnsureAndActivateQueryOutputForTable( _
    ByVal lo As ListObject _
)
    'ensure output exists and navigate the user to it
    If lo Is Nothing Then Exit Sub

    Dim wb As Workbook: Set wb = lo.Parent.Parent
    Dim qName As String: qName = mod_String.SanitizeName(lo.name)

    ' Make sure query + at least one bound sheet exist
    EnsureQueryAndSheetForTable lo

    ' Find first bound ListObject and activate it
    Dim boundLo As ListObject
    Set boundLo = GetFirstBoundListObject(wb, qName)

    If boundLo Is Nothing Then
        ' Very defensive: if for some reason we still didn't find it, create one and retry
        LoadQueryToNewSheet wb, qName
        Set boundLo = GetFirstBoundListObject(wb, qName)
    End If

    If Not boundLo Is Nothing Then
        Application.GoTo boundLo.Range.Cells(1, 1), True  ' scroll focus into view
        boundLo.Parent.Activate
        boundLo.Range.Select
    End If
End Sub

Function UniqueSheetName( _
    ByVal wb As Workbook, _
    ByVal baseName As String _
) As String
    ' Returns a unique worksheet name by appending numeric suffixes if needed.
    Dim uniqueName As String
    Dim suffix As Long
    uniqueName = baseName
    Do While SheetExists(wb, uniqueName)
        suffix = suffix + 1
        uniqueName = baseName & "_" & suffix
    Loop
    UniqueSheetName = uniqueName
End Function

Function SheetExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String _
) As Boolean
    ' True if a worksheet with the given name exists.
    On Error Resume Next
    SheetExists = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
