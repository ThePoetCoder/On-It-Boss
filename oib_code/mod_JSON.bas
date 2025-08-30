Attribute VB_Name = "mod_JSON"
Option Explicit
Option Private Module

Function JSON_GetString( _
    ByVal objJson As String, _
    ByVal keyName As String _
) As String
    ' Returns the unescaped JSON string value for a given key from a minified JSON object.
    Dim s As String
    s = Trim$(objJson)
    If Len(s) = 0 Or Len(keyName) = 0 Then Exit Function

    ' 1) If the WHOLE payload is a JSON *string* (wrapped in quotes), unescape it to get the inner object text.
    '    Do NOT unescape raw JSON objects, or we will turn valid \" into " and break parsing.
    If Len(s) >= 2 Then
        If Left$(s, 1) = mod_String.QUOTE_CHAR And Right$(s, 1) = mod_String.QUOTE_CHAR Then
            On Error Resume Next
            s = mod_JSON.JSON_Unescape(Mid$(s, 2, Len(s) - 2))
            On Error GoTo 0
        End If
    End If

    ' 2) Build a regex that matches:  "keyName"  :  "value"
    '    capturing the JSON-string value including escaped characters.
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.MultiLine = False
    re.IgnoreCase = True  ' JSON keys are technically case-sensitive, but you used vbTextCompare before.

    ' Pattern explanation (PCRE-ish):
    '   "(?:KEY)"\s*:\s*"( (?:[^"\\]|\\.)* )"
    ' Group 1 captures all characters of the JSON string value, allowing for escapes (\" \\ \n \uXXXX etc.)
    re.Pattern = """" & EscapeRegex(keyName) & """\s*:\s*""((?:[^""\\]|\\.)*)"""

    Dim matches As Object
    Set matches = re.Execute(s)
    If matches.Count > 0 Then
        Dim rawValue As String
        rawValue = matches(0).SubMatches(0)  ' the captured, still-escaped JSON string
        JSON_GetString = mod_JSON.JSON_Unescape(rawValue)
    End If
End Function

Function JSON_Unescape( _
    ByVal s As String _
) As String
    ' Unescapes JSON string content including \" \\ \n \r \t and minimal \uXXXX handling.
    Dim r As String, i As Long, ch As String, nxt As String
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = mod_String.BACKSLASH_CHAR Then
            If i = Len(s) Then
                ' Trailing backslash; keep it
                r = r & mod_String.BACKSLASH_CHAR
            Else
                nxt = Mid$(s, i + 1, 1)
                Select Case nxt
                    Case mod_String.QUOTE_CHAR
                        r = r & mod_String.QUOTE_CHAR
                    Case mod_String.BACKSLASH_CHAR
                        r = r & mod_String.BACKSLASH_CHAR
                    Case "/"
                        r = r & "/"
                    Case "b"
                        r = r & vbBack
                    Case "f"
                        r = r & vbFormFeed
                    Case "n"
                        r = r & vbLf
                    Case "r"
                        r = r & vbCr
                    Case "t"
                        r = r & vbTab
                    Case "u"
                        ' Minimal \uXXXX handling
                        If i + 5 <= Len(s) Then
                            Dim hex4 As String, code As Long
                            hex4 = Mid$(s, i + 2, 4)
                            If IsHex4(hex4) Then
                                code = CLng("&H" & hex4)
                                r = r & ChrW(code)
                                i = i + 4 ' extra advance is added below
                            Else
                                r = r & mod_String.BACKSLASH_CHAR & "u" ' keep literal if malformed
                            End If
                        Else
                            r = r & mod_String.BACKSLASH_CHAR & "u"
                        End If
                    Case Else
                        ' Unknown escape, keep the escaped char as-is (tolerant)
                        r = r & nxt
                End Select
                i = i + 1 ' skip the escaped char
            End If
        Else
            r = r & ch
        End If
        i = i + 1
    Loop
    JSON_Unescape = r
End Function

Function LooksLikeEscapedJson( _
    ByVal s As String _
) As Boolean
    ' Heuristically detects if a JSON object is double-escaped OR just raw JSON.
    ' Liberal on purpose: unescaping a raw JSON string is a no-op if there are no escapes.
    Dim bsq As String    ' Backslash + Quote
    Dim bsob As String   ' Backslash + {
    Dim bscb As String   ' Backslash + }
    Dim bscolon As String ' Backslash + ":""

    bsq = mod_String.ESC_QUOTE_CHAR
    bsob = bsq & "{"
    bscb = bsq & "}"
    bscolon = bsq & ":" & mod_String.QUOTE_CHAR

    LooksLikeEscapedJson = (InStr(s, "{") > 0) _
                          Or (InStr(s, bsob) > 0) _
                          Or (InStr(s, bscb) > 0) _
                          Or (InStr(s, bscolon) > 0) _
                          Or (InStr(s, mod_String.BACKSLASH_CHAR & mod_String.BACKSLASH_CHAR) > 0 And InStr(s, bsq) > 0)
End Function

Function EscapeRegex( _
    ByVal text As String _
) As String
    ' Escapes regex metacharacters so a JSON key can be safely used inside VBScript.RegExp.
    ' Metas in VBScript.RegExp: \ ^ $ . | ? * + ( ) [ ] { }
    Dim t As String
    t = text
    t = Replace(t, "\", "\\")
    t = Replace(t, "^", "\^")
    t = Replace(t, "$", "\$")
    t = Replace(t, ".", "\.")
    t = Replace(t, "|", "\|")
    t = Replace(t, "?", "\?")
    t = Replace(t, "*", "\*")
    t = Replace(t, "+", "\+")
    t = Replace(t, "(", "\(")
    t = Replace(t, ")", "\)")
    t = Replace(t, "[", "\[")
    t = Replace(t, "]", "\]")
    t = Replace(t, "{", "\{")
    t = Replace(t, "}", "\}")
    EscapeRegex = t
End Function

Function IsHex4( _
    ByVal s As String _
) As Boolean
    ' True if the 4-character string is valid hex (for \uXXXX parsing).
    Dim i As Long, ch As String
    If Len(s) <> 4 Then Exit Function
    For i = 1 To 4
        ch = Mid$(s, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsHex4 = True
End Function

Function FindMatchingBrace( _
    ByVal text As String, _
    ByVal startPos As Long _
) As Long
    ' Finds the matching closing brace for a JSON object, skipping over strings and escapes.
    Dim braceDepth As Long
    Dim index As Long
    Dim ch As String
    Dim inString As Boolean

    braceDepth = 0
    For index = startPos To Len(text)
        ch = Mid$(text, index, 1)
        If ch = mod_String.QUOTE_CHAR Then
            If index = 1 Or Mid$(text, index - 1, 1) <> mod_String.BACKSLASH_CHAR Then
                inString = Not inString
            End If
        ElseIf Not inString Then
            If ch = "{" Then
                braceDepth = braceDepth + 1
            ElseIf ch = "}" Then
                braceDepth = braceDepth - 1
                If braceDepth = 0 Then
                    FindMatchingBrace = index
                    Exit Function
                End If
            End If
        End If
    Next index
End Function
