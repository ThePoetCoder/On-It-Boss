Attribute VB_Name = "mod_Env"
Option Explicit
Option Private Module

Public TEST_OVERRIDE_EnvPath As String

Function ReadApiKeyFromEnv() As String
    ' Reads OPENAI_API_KEY from the add-in .env file; returns empty string if missing.
    Dim currentLine As String
    Dim equalsPos As Long
    Dim fileNum As Integer
    Dim envPath As String

    fileNum = FreeFile
    envPath = GetAddInsEnvPath()
    On Error GoTo Done
    Open envPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, currentLine
        currentLine = Trim$(currentLine)
        If Left$(currentLine, 1) = "#" Or Len(currentLine) = 0 Then GoTo ContinueLoop
        equalsPos = InStr(1, currentLine, "=", vbTextCompare)
        If equalsPos > 0 Then
            If Trim$(Left$(currentLine, equalsPos - 1)) = "OPENAI_API_KEY" Then
                ReadApiKeyFromEnv = Trim$(Mid$(currentLine, equalsPos + 1))
                Exit Do
            End If
        End If
ContinueLoop:
    Loop
Done:
    On Error Resume Next: Close #fileNum: On Error GoTo 0
End Function

Function GetAddInsEnvPath() As String
    ' Returns the absolute path to the AddIns .env file under the current Windows user profile.
    If Len(TEST_OVERRIDE_EnvPath) > 0 Then
        GetAddInsEnvPath = TEST_OVERRIDE_EnvPath
        Exit Function
    End If
    GetAddInsEnvPath = Environ$("APPDATA") & "\Microsoft\AddIns\.env"
End Function

Public Sub WriteApiKeyToEnv(ByVal newKey As String)
    ' Updates or inserts OPENAI_API_KEY in the .env file.
    Dim lines As Collection
    Set lines = New Collection
    
    Dim fileNum As Integer
    Dim currentLine As String
    Dim equalsPos As Long
    Dim found As Boolean
    Dim envPath As String
    
    envPath = GetAddInsEnvPath()
    fileNum = FreeFile
    
    ' Read existing .env (if present)
    On Error Resume Next
    Open envPath For Input As #fileNum
    If Err.Number = 0 Then
        Do While Not EOF(fileNum)
            Line Input #fileNum, currentLine
            equalsPos = InStr(1, currentLine, "=", vbTextCompare)
            If equalsPos > 0 Then
                If Trim$(Left$(currentLine, equalsPos - 1)) = "OPENAI_API_KEY" Then
                    lines.Add "OPENAI_API_KEY=" & newKey
                    found = True
                Else
                    lines.Add currentLine
                End If
            Else
                lines.Add currentLine
            End If
        Loop
        Close #fileNum
    Else
        Err.clear
    End If
    On Error GoTo 0
    
    ' If no existing line found, append it
    If Not found Then
        lines.Add "OPENAI_API_KEY=" & newKey
    End If
    
    ' Write back file
    fileNum = FreeFile
    Open envPath For Output As #fileNum
    Dim i As Long
    For i = 1 To lines.Count
        Print #fileNum, lines(i)
    Next
    Close #fileNum
End Sub

