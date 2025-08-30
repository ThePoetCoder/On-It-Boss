Attribute VB_Name = "mod_API"
Option Explicit
Option Private Module

Public TEST_OVERRIDE_ContentJson As String

Function CallChatAPI_JSON( _
    apiKey As String, _
    payload As String _
) As String
    ' Performs the HTTP POST to OpenAI with retries/backoff; returns the inner content JSON string.
    If Len(TEST_OVERRIDE_ContentJson) > 0 Then
        CallChatAPI_JSON = TEST_OVERRIDE_ContentJson
        TEST_OVERRIDE_ContentJson = ""
        Exit Function
    End If
    
    Dim http As Object
    On Error GoTo Fail
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", OPENAI_CHAT_ENDPOINT, False
    
    ' Resolve, Connect, Send, Receive (ms)
    http.SetTimeouts 30000, 30000, 30000, 180000
    SetHeader http, apiKey
    
    Dim attempts As Integer: attempts = 0
    Do
        attempts = attempts + 1
        mod_Core.Panel_Status "Calling model (attempt " & attempts & "/3)…"
        
        http.Send payload
        If http.Status = 200 Then
            Dim responseBody As String
            responseBody = http.ResponseText
            CallChatAPI_JSON = ExtractFirstMessageContent_JSON(responseBody)
            Exit Function
        End If
        If (http.Status = 429 Or (http.Status >= 500 And http.Status < 600)) And attempts < 3 Then
            Dim waitMs As Long: waitMs = 500 * attempts
            Dim t As Single: t = Timer: Do While Timer < t + (waitMs / 1000!): DoEvents: Loop
            http.Open "POST", OPENAI_CHAT_ENDPOINT, False
            SetHeader http, apiKey
        Else
            MsgBox "OpenAI API error " & http.Status & ": " & http.ResponseText, vbExclamation
            CallChatAPI_JSON = vbNullString
            Exit Function
        End If
    Loop

Fail:
    MsgBox "HTTP failure calling OpenAI: " & Err.Description, vbCritical
    CallChatAPI_JSON = vbNullString
End Function

Sub SetHeader( _
    ByRef http As Object, _
    apiKey As String _
)
    'Sets http headers
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
End Sub

Function ExtractFirstMessageContent_JSON( _
    ByVal fullJson As String _
) As String
    ' Extracts choices[0].message.content which may be:
    '  (a) a JSON object    -> {...}
    '  (b) a JSON string    -> " {...} "  (escaped JSON)
    ' Returns the *JSON object string* in both cases.

    Dim searchToken As String
    Dim contentKeyPos As Long
    Dim i As Long
    Dim ch As String

    searchToken = mod_String.QUOTE_CHAR & "content" & mod_String.QUOTE_CHAR & ":"
    contentKeyPos = InStr(1, fullJson, searchToken, vbTextCompare)
    If contentKeyPos = 0 Then Exit Function

    ' Advance to the first non-whitespace char after the colon
    i = contentKeyPos + Len(searchToken)
    Do While i <= Len(fullJson)
        ch = Mid$(fullJson, i, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        i = i + 1
    Loop
    If i > Len(fullJson) Then Exit Function

    If ch = mod_String.QUOTE_CHAR Then
        ' Case (b): content is a JSON STRING -> extract the string token safely, then unescape
        Dim startStr As Long, j As Long, prev As String, inEsc As Boolean
        startStr = i + 1 ' first char after opening quote
        j = startStr
        inEsc = False
        Do While j <= Len(fullJson)
            ch = Mid$(fullJson, j, 1)
            If inEsc Then
                inEsc = False
            ElseIf ch = mod_String.BACKSLASH_CHAR Then
                inEsc = True
            ElseIf ch = mod_String.QUOTE_CHAR Then
                Exit Do ' closing quote
            End If
            j = j + 1
        Loop
        If j > Len(fullJson) Then Exit Function
        ' Extract the raw (escaped) JSON string payload
        Dim rawEscaped As String
        rawEscaped = Mid$(fullJson, startStr, j - startStr)
        ' Unescape it to get the actual JSON object text
        ExtractFirstMessageContent_JSON = mod_JSON.JSON_Unescape(rawEscaped)
        Exit Function

    ElseIf ch = "{" Then
        ' Case (a): content is a JSON OBJECT -> find matching brace
        Dim contentEndBracePos As Long
        contentEndBracePos = mod_JSON.FindMatchingBrace(fullJson, i)
        If contentEndBracePos > i Then
            ExtractFirstMessageContent_JSON = Mid$(fullJson, i, contentEndBracePos - i + 1)
            Exit Function
        End If
    End If
    ' Otherwise (unexpected shape), return empty.
End Function

