Attribute VB_Name = "mod_String"
Option Explicit
Option Private Module

Public Const QUOTE_CHAR As String = """"                  ' a single double-quote character"
Public Const BACKSLASH_CHAR As String = "\"               ' a single backslash character
Public Const ESC_QUOTE_CHAR = BACKSLASH_CHAR & QUOTE_CHAR ' combination of the above 2

Function EscapeJson( _
    ByVal s As String _
) As String
    ' Escapes a string for JSON (quotes, backslashes, and newlines).
    s = Replace(s, BACKSLASH_CHAR, BACKSLASH_CHAR & BACKSLASH_CHAR)
    s = Replace(s, QUOTE_CHAR, ESC_QUOTE_CHAR)
    s = Replace(s, vbCrLf, BACKSLASH_CHAR & "n")
    s = Replace(s, vbCr, BACKSLASH_CHAR & "n")
    s = Replace(s, vbLf, BACKSLASH_CHAR & "n")
    s = Replace(s, vbTab, BACKSLASH_CHAR & "t")
    EscapeJson = s
End Function

Function SanitizeName( _
    ByVal s As String _
) As String
    ' Produces a safe query name from an arbitrary table name (spaces?underscore, slashes?dash).
    Dim t As String
    t = s
    t = Replace(t, " ", "_")
    t = Replace(t, "/", "-")
    t = Replace(t, "\", "-")
    SanitizeName = t
End Function

Function SanitizeTableName( _
    ByVal qName As String _
) As String
    ' Makes a ListObject-safe name from a query name (letters/numbers/underscore).
    Dim s As String
    s = qName
    s = Replace(s, " ", "_")
    s = Replace(s, "-", "_")
    s = Replace(s, ".", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, BACKSLASH_CHAR, "_")
    s = Replace(s, "(", "_")
    s = Replace(s, ")", "_")
    If Not s Like "[A-Za-z]*" Then s = "T_" & s
    If Len(s) > 240 Then s = Left$(s, 240) ' leave room for "tbl_"
    SanitizeTableName = s
End Function

Function NormalizeModelCode( _
    ByVal s As String _
) As String
    ' Normalizes model-returned M: unescape layers, normalize newlines, strip fences/quotes/BOM.
    Dim i As Long
    Dim before As String
    Dim shouldUnescape As Boolean

    ' Defensive: if the entire payload came back double-escaped, peel it a few times.
    For i = 1 To 3
        before = s
        shouldUnescape = False

        ' Only unescape when the WHOLE blob appears to be a single JSON string
        ' (wrapped in quotes) that still contains JSON-style escapes.
        If Len(s) >= 2 Then
            If Left$(s, 1) = QUOTE_CHAR And Right$(s, 1) = QUOTE_CHAR Then
                If InStr(1, s, ESC_QUOTE_CHAR, vbBinaryCompare) > 0 _
                   Or InStr(1, s, BACKSLASH_CHAR & "n", vbBinaryCompare) > 0 _
                   Or InStr(1, s, BACKSLASH_CHAR & BACKSLASH_CHAR, vbBinaryCompare) > 0 _
                   Or InStr(1, s, BACKSLASH_CHAR & "u", vbBinaryCompare) > 0 Then
                    shouldUnescape = True
                End If
            End If
        End If

        If shouldUnescape Then
            s = mod_JSON.JSON_Unescape(s)
        Else
            Exit For
        End If

        If s = before Then Exit For
    Next i

    ' Normalize newlines (fix: collapse CRLF first, then lone CR)
    s = Replace(s, vbCrLf, vbLf)            ' CRLF -> LF (no duplication)
    s = Replace(s, vbCr, vbLf)              ' bare CR -> LF
    s = Replace(s, "\r\n", vbLf)            ' literal \r\n (if any survived)
    s = Replace(s, "\n", vbLf)              ' literal \n (if any survived)
    s = Replace(s, vbLf, vbCrLf)            ' final: LF -> CRLF

    ' Strip Markdown/code fences if the model ever returns them
    s = Trim$(s)
    If Left$(s, 3) = "```" Then
        Dim fenceEnd As Long, lfPos As Long, newlineLen As Long
        fenceEnd = InStr(4, s, vbCrLf)
        newlineLen = 2
        If fenceEnd = 0 Then
            lfPos = InStr(4, s, vbLf)
            If lfPos > 0 Then
                fenceEnd = lfPos
                newlineLen = 1
            End If
        End If
        If fenceEnd > 0 Then
            s = Mid$(s, fenceEnd + newlineLen)
        Else
            Dim secondFence As Long
            secondFence = InStr(4, s, "```")
            If secondFence > 0 Then s = Mid$(s, secondFence + 3)
        End If

        Dim closePos As Long
        closePos = InStrRev(s, "```")
        If closePos > 0 Then s = Left$(s, closePos - 1)

        s = Trim$(s)
    End If

    ' Remove any leading BOM (must be before quote stripping!)
    If Len(s) > 0 And AscW(Left$(s, 1)) = &HFEFF Then
        s = Mid$(s, 2)
    End If

    ' Strip wrapping quotes if the whole blob is still quoted
    If Len(s) >= 2 Then
        If Left$(s, 1) = QUOTE_CHAR And Right$(s, 1) = QUOTE_CHAR Then
            s = Mid$(s, 2, Len(s) - 2)
        End If
    End If
    
    If InStr(1, s, "Table.Combine(", vbTextCompare) > 0 Then
        s = FixTableCombineSyntax(s)
    End If

    NormalizeModelCode = s
End Function

Public Function FixTableCombineSyntax(ByVal m As String) As String
    Dim i As Long, n As Long
    Dim res As String
    Dim last As Long: last = 1
    Dim inString As Boolean, inSL As Boolean, inML As Boolean
    Dim ch As String, ch2 As String
    
    n = Len(m)
    i = 1
    Do While i <= n
        ch = Mid$(m, i, 1)
        If i < n Then ch2 = Mid$(m, i + 1, 1) Else ch2 = vbNullString
        
        ' State machine: strings/comments
        If inString Then
            If ch = """" Then
                If ch2 = """" Then
                    i = i + 2
                Else
                    inString = False: i = i + 1
                End If
            Else
                i = i + 1
            End If
            GoTo ContinueLoop
        ElseIf inSL Then
            If ch = vbCr Or ch = vbLf Then inSL = False
            i = i + 1
            GoTo ContinueLoop
        ElseIf inML Then
            If ch = "*" And ch2 = "/" Then
                inML = False: i = i + 2
            Else
                i = i + 1
            End If
            GoTo ContinueLoop
        Else
            If ch = """" Then
                inString = True: i = i + 1: GoTo ContinueLoop
            ElseIf ch = "/" And ch2 = "/" Then
                inSL = True: i = i + 2: GoTo ContinueLoop
            ElseIf ch = "/" And ch2 = "*" Then
                inML = True: i = i + 2: GoTo ContinueLoop
            End If
        End If
        
        ' Detect "Table.Combine" (case-insensitive)
        If i + 12 <= n Then
            If StrComp(Mid$(m, i, 13), "Table.Combine", vbTextCompare) = 0 Then
                Dim j As Long: j = i + 13
                ' Skip whitespace before '('
                Do While j <= n And IsWs(Mid$(m, j, 1))
                    j = j + 1
                Loop
                If j <= n And Mid$(m, j, 1) = "(" Then
                    Dim startP As Long: startP = j
                    Dim endP As Long: endP = FindMatchingParen(m, startP + 1)
                    If endP > 0 Then
                        ' If first non-ws inside is "{", it's already a list: leave unchanged
                        Dim p As Long: p = FirstNonWs(m, startP + 1, endP - 1)
                        If p > 0 And Mid$(m, p, 1) = "{" Then
                            i = endP + 1
                            GoTo ContinueLoop
                        Else
                            ' --- FIX: remove any spaces before "(" when rewriting ---
                            Dim nameEnd As Long: nameEnd = i + 13 - 1 ' end of "Table.Combine"
                            Dim inner As String
                            inner = Mid$(m, startP + 1, endP - startP - 1)
                            inner = TrimAll(inner)
                            
                            ' Copy up to the end of the name (no pre-paren spaces),
                            ' then insert "(" + "{" + inner + "}" + ")"
                            res = res & Mid$(m, last, nameEnd - last + 1)
                            res = res & "(" & "{" & inner & "}" & ")"
                            
                            last = endP + 1
                            i = endP + 1
                            GoTo ContinueLoop
                        End If
                    End If
                End If
            End If
        End If
        
        i = i + 1
ContinueLoop:
    Loop
    
    If last <= n Then res = res & Mid$(m, last, n - last + 1)
    FixTableCombineSyntax = res
End Function

'=== Helpers ===

Private Function IsWs(ByVal s As String) As Boolean
    IsWs = (s = " " Or s = vbTab Or s = vbCr Or s = vbLf)
End Function

Private Function FirstNonWs(ByVal s As String, ByVal lo As Long, ByVal hi As Long) As Long
    Dim k As Long
    For k = lo To hi
        If Not IsWs(Mid$(s, k, 1)) Then
            FirstNonWs = k
            Exit Function
        End If
    Next k
    FirstNonWs = 0
End Function

Private Function TrimAll(ByVal s As String) As String
    Dim a As Long, b As Long, L As Long
    L = Len(s): a = 1: b = L
    Do While a <= b And IsWs(Mid$(s, a, 1)): a = a + 1: Loop
    Do While b >= a And IsWs(Mid$(s, b, 1)): b = b - 1: Loop
    If b >= a Then TrimAll = Mid$(s, a, b - a + 1) Else TrimAll = vbNullString
End Function

Private Function FindMatchingParen(ByVal s As String, ByVal startPos As Long) As Long
    ' s(startPos-1) must be "("; find the index of its matching ")"
    Dim n As Long: n = Len(s)
    Dim i As Long: i = startPos
    Dim depth As Long: depth = 1
    Dim inString As Boolean, inSL As Boolean, inML As Boolean
    Dim ch As String, ch2 As String
    
    Do While i <= n
        ch = Mid$(s, i, 1)
        If i < n Then ch2 = Mid$(s, i + 1, 1) Else ch2 = vbNullString
        
        If inString Then
            If ch = """" Then
                If ch2 = """" Then
                    i = i + 2
                Else
                    inString = False
                    i = i + 1
                End If
            Else
                i = i + 1
            End If
        ElseIf inSL Then
            If ch = vbCr Or ch = vbLf Then inSL = False
            i = i + 1
        ElseIf inML Then
            If ch = "*" And ch2 = "/" Then
                inML = False
                i = i + 2
            Else
                i = i + 1
            End If
        Else
            If ch = """" Then
                inString = True: i = i + 1
            ElseIf ch = "/" And ch2 = "/" Then
                inSL = True: i = i + 2
            ElseIf ch = "/" And ch2 = "*" Then
                inML = True: i = i + 2
            ElseIf ch = "(" Then
                depth = depth + 1: i = i + 1
            ElseIf ch = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    FindMatchingParen = i
                    Exit Function
                End If
                i = i + 1
            Else
                i = i + 1
            End If
        End If
    Loop
    
    FindMatchingParen = 0 ' not found
End Function
