Attribute VB_Name = "mod_Core"
Option Explicit
Option Private Module

Private gPrevPanelCaption As String
Private gIsPanelBusy As Boolean

Sub ChatOIB( _
    ByVal promptText As String, _
    ByVal lo As ListObject _
)
    ' Entrypoint from the UI: validates input and runs a transform for the active ListObject.
    Dim userPrompt As String: userPrompt = Trim$(promptText)
    
    If Len(userPrompt) = 0 Then
        MsgBox "Type something to transform.", vbExclamation
        Exit Sub
    End If
    
    Call RunTransformForListObject(lo, userPrompt, True)
End Sub

Function RunTransformForListObject( _
    ByVal lo As ListObject, _
    ByVal userPrompt As String, _
    ByVal saveHistory As Boolean _
) As Boolean
    ' Core pipeline that builds prompt, calls the model, applies M, and optionally records history.
    If lo Is Nothing Then
        MsgBox "No table selected.", vbExclamation
        Exit Function
    End If
    userPrompt = Trim$(userPrompt)
    If Len(userPrompt) = 0 Then
        MsgBox "Type something to transform.", vbExclamation
        Exit Function
    End If
    
    Panel_Status "Preparing schema…"

    Dim originSchemaJson As String
    Dim currentSchemaJson As String
    Dim currentM As String
    Dim queryName As String
    BuildSchemaAndCurrentM lo, originSchemaJson, currentSchemaJson, currentM, queryName
    
    Dim xmlPrompt As String
    xmlPrompt = BuildXmlPrompt(currentM, originSchemaJson, currentSchemaJson, userPrompt)

    mod_History.History_EnsureSeedForTable lo.Parent.Parent, lo

    Dim retTitle As String, retLang As String, retM As String
    If Not RunChatAndParse(xmlPrompt, retTitle, retLang, retM) Then Exit Function

    ' Apply immediately (keeps identical behavior via existing helper)
    mod_SchemaQuery.ApplyMToTable lo, retM, retTitle

    ' History: only if requested (preserves current difference between the two entrypoints)
    If saveHistory Then
        mod_History.History_Add lo.Parent.Parent, lo.name, queryName, retTitle, retLang, retM
    End If
    
    Panel_StatusDoneAndClear

    RunTransformForListObject = True
End Function

Function RunChatAndParse( _
    ByVal xmlPrompt As String, _
    ByRef outTitle As String, _
    ByRef outLang As String, _
    ByRef outM As String _
) As Boolean
    ' Calls the model and parses the strict JSON into title/language/M code; enforces language safety.
    Dim apiKey As String: apiKey = mod_Env.ReadApiKeyFromEnv()
    If Len(apiKey) = 0 Then
        MsgBox "API key not found. Create a .env in your AddIns folder with OPENAI_API_KEY=...", vbCritical
        Exit Function
    End If

    Dim modelName As String, temp As Double
    modelName = uf_Panel.SelectedModel
    temp = uf_Panel.SelectedTemperature
    If mod_Config.DEBUG_CHAT Then Debug.Print "Model: '" & modelName & " || 'Temp: '" & temp

    Dim payload As String: payload = BuildChatJSONPayload(modelName, temp, xmlPrompt)
    If mod_Config.DEBUG_CHAT Then Debug.Print "Payload:" & vbCrLf & payload

    Dim contentJson As String
    contentJson = mod_API.CallChatAPI_JSON(apiKey, payload)
    If mod_Config.DEBUG_CHAT Then Debug.Print "Response:" & vbCrLf & contentJson & vbCrLf
    If Len(contentJson) = 0 Then Exit Function

    outTitle = mod_JSON.JSON_GetString(contentJson, "title")
    outLang = LCase$(mod_JSON.JSON_GetString(contentJson, "language"))
    outM = mod_String.NormalizeModelCode(mod_JSON.JSON_GetString(contentJson, "code"))

    If outLang <> "m" Then
        MsgBox "Response was not language='m'. Aborting for safety.", vbExclamation
        Exit Function
    End If
    If Len(outM) = 0 Then
        MsgBox "No M code returned.", vbExclamation
        Exit Function
    End If

    RunChatAndParse = True
End Function

Function BuildXmlPrompt( _
    ByVal currentM As String, _
    ByVal originSchemaJson As String, _
    ByVal currentSchemaJson As String, _
    ByVal userRequest As String _
) As String
    ' Builds the XML-style prompt with <CURRENT_M_CODE>, <ORIGIN_TABLE_SCHEMA>, <TABLE_SCHEMA_SAMPLE>, <USER_REQUEST>.
    Dim xml As String
    xml = "<GUIDELINES>" & vbCrLf & _
          " * Use <CURRENT_M_CODE> as a starting point for your response and build upon it:" & vbCrLf & _
          "     * Use the <CURRENT_M_CODE> to intuit what the user requested previously and <USER_REQUEST> to determine what they are asking to change." & vbCrLf & _
          "     * Do not remove code unless you are specifically requested to do so." & vbCrLf & _
          "     * Assume that you will copy and paste what is given in the <CURRENT_M_CODE> and only add steps AFTER that unless you are specifically requested change a previous step." & vbCrLf & _
          " * Treat <ORIGIN_TABLE_SCHEMA> as the schema at the `Source` step in <CURRENT_M_CODE> and trace the lineage of transformations from there to <CURRENT_TABLE_SCHEMA>." & vbCrLf & _
          " * DO NOT modify the Source step, especially if <ORIGIN_TABLE_SCHEMA> is empty ({}) or unavailable in which case you should infer available columns from <CURRENT_TABLE_SCHEMA> and append new steps after the existing ones." & vbCrLf & _
          " * If a column is a 'Date' column, make sure it is formatted that way." & vbCrLf & _
          " * Format and indent the M code in your response according to best practices and lean towards making the fewest changes to <CURRENT_M_CODE> as possible." & vbCrLf & _
          "</GUIDELINES>" & vbCrLf & _
          "<CURRENT_M_CODE>" & vbCrLf & _
          currentM & vbCrLf & _
          "</CURRENT_M_CODE>" & vbCrLf & _
          "<ORIGIN_TABLE_SCHEMA>" & vbCrLf & _
          originSchemaJson & vbCrLf & _
          "</ORIGIN_TABLE_SCHEMA>" & vbCrLf & _
          "<CURRENT_TABLE_SCHEMA>" & vbCrLf & _
          currentSchemaJson & vbCrLf & _
          "</CURRENT_TABLE_SCHEMA>" & vbCrLf & _
          "<USER_REQUEST>" & vbCrLf & _
          userRequest & vbCrLf & _
          "</USER_REQUEST>"
    BuildXmlPrompt = xml
End Function

Function BuildChatJSONPayload( _
    modelName As String, _
    temp As Double, _
    xmlPrompt As String _
) As String
    ' Builds the Chat Completions JSON payload with system and user messages and json_object format.
    Dim systemMsg As String
    Dim userMsg As String

    systemMsg = _
        "You generate Power Query M code ONLY." & vbCrLf & _
        "You will be given an XML-style prompt with these tags:" & vbCrLf & _
        "<CURRENT_M_CODE>, <ORIGIN_TABLE_SCHEMA>, <CURRENT_TABLE_SCHEMA>, <TABLE_SCHEMA_SAMPLE>, <USER_REQUEST>." & vbCrLf & _
        "Start from <CURRENT_M_CODE>; make the minimal changes needed." & vbCrLf & _
        "Treat <ORIGIN_TABLE_SCHEMA> as the columns available at the `Source` step." & vbCrLf & _
        "Treat <CURRENT_TABLE_SCHEMA> as the columns produced by CURRENT_M_CODE so far." & vbCrLf & _
        "If the user requests logic that uses a column that does not yet exist, you MUST add an explicit step to create it BEFORE referencing it (e.g., before Table.Group, filters, or joins)." & vbCrLf & _
        "Return a strict MINIFIED JSON object with keys: title (string), language ('m'), code (string)." & vbCrLf & _
        "The `code` value MUST be wrapped in triple backticks, like " & mod_String.QUOTE_CHAR & "```m ... ```" & mod_String.QUOTE_CHAR & "." & vbCrLf & _
        "No explanations, no extra markdown outside of the JSON. The code must be valid, self-contained M."

    userMsg = xmlPrompt  ' already built and escaped

    BuildChatJSONPayload = "{" & _
        mod_String.QUOTE_CHAR & "model" & mod_String.QUOTE_CHAR & ":" & mod_String.QUOTE_CHAR & mod_String.EscapeJson(modelName) & mod_String.QUOTE_CHAR & "," & _
        mod_String.QUOTE_CHAR & "temperature" & mod_String.QUOTE_CHAR & ":" & CStr(temp) & "," & _
        mod_String.QUOTE_CHAR & "response_format" & mod_String.QUOTE_CHAR & ":{" & mod_String.QUOTE_CHAR & "type" & mod_String.QUOTE_CHAR & ":" & mod_String.QUOTE_CHAR & "json_object" & mod_String.QUOTE_CHAR & "}," & _
        mod_String.QUOTE_CHAR & "messages" & mod_String.QUOTE_CHAR & ":[" & _
            "{" & mod_String.QUOTE_CHAR & "role" & mod_String.QUOTE_CHAR & ":" & mod_String.QUOTE_CHAR & "system" & mod_String.QUOTE_CHAR & "," & mod_String.QUOTE_CHAR & "content" & mod_String.QUOTE_CHAR & ":" & mod_String.QUOTE_CHAR & mod_String.EscapeJson(systemMsg) & mod_String.QUOTE_CHAR & "}," & _
            "{" & mod_String.QUOTE_CHAR & "role" & mod_String.QUOTE_CHAR & ":" & mod_String.QUOTE_CHAR & "user" & mod_String.QUOTE_CHAR & "," & mod_String.QUOTE_CHAR & "content" & mod_String.QUOTE_CHAR & ":" & mod_String.QUOTE_CHAR & mod_String.EscapeJson(userMsg) & mod_String.QUOTE_CHAR & "}" & _
        "]" & _
    "}"
End Function

Sub BuildSchemaAndCurrentM( _
    ByVal lo As ListObject, _
    ByRef outOriginSchemaJson As String, _
    ByRef outCurrentSchemaJson As String, _
    ByRef outCurrentM As String, _
    ByRef outQueryName As String _
)
    ' Produces ORIGIN + CURRENT schema JSON, current M, and sanitized query name.
    Dim wb As Workbook
    Set wb = lo.Parent.Parent

    outQueryName = mod_String.SanitizeName(lo.name)
    outCurrentM = GetOrInitCurrentM(wb, outQueryName, lo.name)

    ' CURRENT schema = the active (bound) ListObject you’re working on today
    outCurrentSchemaJson = mod_SchemaQuery.BuildTableSchemaJson(lo)

    ' ORIGIN schema = the Excel.CurrentWorkbook(){[Name="..."]}[Content] table used in M
    outOriginSchemaJson = mod_SchemaQuery.BuildOriginSchemaJson(wb, outCurrentM)
End Sub


Function GetOrInitCurrentM( _
    ByVal wb As Workbook, _
    ByVal queryName As String, _
    ByVal loName As String _
) As String
    ' Returns existing query M or a minimal Excel.CurrentWorkbook() stub if none exists.
    Dim currentM As String
    currentM = mod_SchemaQuery.GetExistingQueryFormula(wb, queryName)
    If Len(Trim$(currentM)) = 0 Then
        currentM = mod_Core.BuildSeedM(loName)
    End If
    GetOrInitCurrentM = currentM
End Function

Sub Panel_Status( _
    ByVal msg As String, _
    Optional ByVal clear As Boolean = False _
)
    ' Sets the panel caption to show progress, preserving/restoring the prior caption safely.
    On Error Resume Next
    If clear Then
        If Len(gPrevPanelCaption) > 0 Then uf_Panel.Caption = gPrevPanelCaption
        gPrevPanelCaption = vbNullString
        gIsPanelBusy = False
    Else
        If Not gIsPanelBusy Then
            gPrevPanelCaption = uf_Panel.Caption
            gIsPanelBusy = True
        End If
        uf_Panel.Caption = mod_Config.title & msg
    End If
    DoEvents ' allow UI repaint
    On Error GoTo 0
End Sub

Sub Panel_StatusDoneAndClear()
    ' Briefly shows “Done!” then restores the panel caption to its previous value.
    Panel_Status "Done!"
    DoEvents
    On Error Resume Next
    Application.Wait Now + TimeSerial(0, 0, 2) 'pause
    On Error GoTo 0
    Panel_Status "", True
End Sub

Function BuildSeedM( _
    ByVal loName As String _
) As String
    ' Returns the minimal Excel.CurrentWorkbook() M for a given table name.
    BuildSeedM = "let" & vbCrLf & _
                 "    Source = Excel.CurrentWorkbook(){[Name=""" & loName & """]}[Content]" & vbCrLf & _
                 "in" & vbCrLf & _
                 "    Source"
End Function


