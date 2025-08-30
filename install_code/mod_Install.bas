Attribute VB_Name = "mod_Install"
'===== mod_InstallerButtons =====================================
Option Explicit

' --- Tweak these as needed ---
Private Const SHEET_NAME As String = "On It, Boss!"
Private Const ADDIN_FILENAME As String = "On It, Boss!.xlam"
Private Const BTN_IDLE_TEXT_COLOR As Long = vbWhite

' Material-ish green & red
Private Const BTN_IDLE_COLOR As Long = &H53& + &H96& * &H100& + &H21& * &H10000   ' RGB(33,150,83)
Private Const BTN_BUSY_COLOR As Long = &H0& + &H0& * &H100& + &HC0& * &H10000     ' RGB(192,0,0)

' ---------- Public entrypoints you assign to the shapes ----------
Public Sub shp_Install_Click()
    RunWithButtonStatus "shp_Install", "Worker_Install", "Install add-in", "Installing…"
End Sub

Public Sub shp_Uninstall_Click()
    RunWithButtonStatus "shp_Uninstall", "Worker_Uninstall", "Uninstall add-in", "Uninstalling…"
End Sub

Private Sub RunWithButtonStatus( _
    ByVal buttonShapeName As String, _
    ByVal workerProcName As String, _
    Optional ByVal idleCaption As String = "Ready", _
    Optional ByVal busyCaption As String = "Working…" _
)
    Dim shp As Shape
    Set shp = GetButtonShape(buttonShapeName)
    If shp Is Nothing Then Exit Sub

    Dim prevText As String, prevColor As Long
    prevText = GetShapeText(shp)
    prevColor = shp.Fill.ForeColor.RGB

    On Error GoTo Fail
    UpdateButtonLook shp, busyCaption, BTN_BUSY_COLOR
    Application.Run workerProcName, buttonShapeName
    UpdateButtonLook shp, idleCaption, BTN_IDLE_COLOR
    Exit Sub

Fail:
    UpdateButtonLook shp, "Failed: " & Err.Description, BTN_BUSY_COLOR
End Sub

Public Sub Worker_Install(ByVal buttonShapeName As String)
    ' 1) Source beside installer
    Dim srcXlam As String
    srcXlam = GetInstallerFolder() & "\" & ADDIN_FILENAME

    PushStatus buttonShapeName, "Locating add-in…"
    If Dir$(srcXlam, vbNormal) = "" Then Err.Raise 53, , "Add-in file not found at: " & srcXlam

    ' 2) Destination = user AddIns folder
    Dim addinsFolder As String, dstXlam As String
    addinsFolder = GetAddInsFolder()
    If Len(Dir$(addinsFolder, vbDirectory)) = 0 Then MkDir addinsFolder
    dstXlam = addinsFolder & "\" & ADDIN_FILENAME

    ' 3) Copy to canonical location (overwrite OK)
    If StrComp(srcXlam, dstXlam, vbTextCompare) <> 0 Then
        PushStatus buttonShapeName, "Copying to AddIns…"
        On Error Resume Next
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile srcXlam, dstXlam, True
        If Err.Number <> 0 Then
            Dim e As String: e = Err.Description: Err.Clear
            On Error GoTo 0
            Err.Raise 75, , "Failed to copy add-in: " & e
        End If
        On Error GoTo 0
    End If

    ' 4) Disable any stale registrations that point to another path
    PushStatus buttonShapeName, "Fixing registration…"
    DisableStaleAddInRegistrations ADDIN_FILENAME, dstXlam

    ' 5) Ensure there's a registration for the canonical path
    Dim add_in As AddIn
    Set add_in = FindAddInByExactPath(dstXlam)
    If add_in Is Nothing Then
        Set add_in = Application.AddIns.Add(Filename:=dstXlam, CopyFile:=False)
    End If

    ' 6) Turn it on
    If Not add_in.Installed Then
        PushStatus buttonShapeName, "Enabling…"
        add_in.Installed = True
    Else
        PushStatus buttonShapeName, "Already enabled"
    End If

    PushStatus buttonShapeName, "Installed"
End Sub

Public Sub Worker_Uninstall(ByVal buttonShapeName As String)
    Dim addinsFolder As String, dstXlam As String
    addinsFolder = GetAddInsFolder()
    dstXlam = addinsFolder & "\" & ADDIN_FILENAME

    PushStatus buttonShapeName, "Finding add-in…"
    Dim add_in As AddIn
    Set add_in = FindAddInByExactPath(dstXlam)
    If add_in Is Nothing Then
        ' Fallback: by name (in case Excel kept a weird entry)
        Dim ai As AddIn
        For Each ai In Application.AddIns
            If StrComp(ai.Name, ADDIN_FILENAME, vbTextCompare) = 0 Then
                Set add_in = ai: Exit For
            End If
        Next ai
    End If
    If add_in Is Nothing Then Err.Raise 5, , "Add-in not registered in Excel."

    If add_in.Installed Then
        PushStatus buttonShapeName, "Disabling…"
        add_in.Installed = False
    Else
        PushStatus buttonShapeName, "Already disabled"
    End If

    ' Tidy any cached entry by name
    On Error Resume Next
    Application.AddIns(add_in.Name).Installed = False
    On Error GoTo 0

    ' Delete the copied file
    If Len(Dir$(dstXlam, vbNormal)) > 0 Then
        PushStatus buttonShapeName, "Deleting file…"
        On Error Resume Next
        Kill dstXlam
        If Err.Number <> 0 Then
            MsgBox "Could not delete: " & dstXlam & vbCrLf & Err.Description, vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    End If

    PushStatus buttonShapeName, "Uninstalled"
End Sub

' ---------- Status / Shape helpers ----------
Public Sub PushStatus(ByVal buttonShapeName As String, ByVal msg As String)
    Dim shp As Shape
    Set shp = GetButtonShape(buttonShapeName)
    If shp Is Nothing Then Exit Sub
    SetShapeText shp, msg
    DoEvents
    Application.Wait (Now + TimeValue("0:00:02"))
End Sub

Private Function GetButtonShape(ByVal buttonShapeName As String) As Shape
    On Error Resume Next
    Set GetButtonShape = ThisWorkbook.Worksheets(SHEET_NAME).Shapes(buttonShapeName)
    If GetButtonShape Is Nothing Then
        MsgBox "Shape '" & buttonShapeName & "' not found on " & SHEET_NAME, vbExclamation
    End If
End Function

Private Sub UpdateButtonLook(ByVal shp As Shape, ByVal caption As String, ByVal fillRgb As Long)
    shp.Fill.Visible = msoTrue
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = fillRgb
    shp.Line.Visible = msoFalse

    SetShapeText shp, caption

    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BTN_IDLE_TEXT_COLOR
    shp.TextFrame.Characters.Font.Color = BTN_IDLE_TEXT_COLOR
    On Error GoTo 0
End Sub

Private Sub SetShapeText(ByVal shp As Shape, ByVal caption As String)
    On Error Resume Next
    If shp.TextFrame2.HasText Then
        shp.TextFrame2.TextRange.Text = caption
    Else
        shp.TextFrame.Characters.Text = caption
    End If
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.AutoSize = msoFalse
    On Error GoTo 0
End Sub

Private Function GetShapeText(ByVal shp As Shape) As String
    On Error Resume Next
    If shp.TextFrame2.HasText Then
        GetShapeText = shp.TextFrame2.TextRange.Text
    Else
        GetShapeText = shp.TextFrame.Characters.Text
    End If
End Function

' ---------- Path / Add-in discovery helpers ----------
Private Function GetInstallerFolder() As String
    ' Folder containing Install.xlsm
    GetInstallerFolder = ThisWorkbook.Path
End Function

Private Function FindAddInByPathOrName(ByVal fullPath As String, ByVal fallbackName As String) As AddIn
    Dim add_in As AddIn
    For Each add_in In Application.AddIns
        If StrComp(add_in.FullName, fullPath, vbTextCompare) = 0 _
           Or StrComp(add_in.Name, fallbackName, vbTextCompare) = 0 _
           Or LCase$(Right$(add_in.FullName, Len(fallbackName))) = LCase$(fallbackName) Then
            Set FindAddInByPathOrName = add_in
            Exit Function
        End If
    Next add_in
End Function

Function GetAddInsFolder() As String
    ' Returns the absolute path to the AddIns folder under the current Windows user profile.
    Dim userName As String
    userName = Environ$("USERNAME")
    GetAddInsFolder = "C:\Users\" & userName & "\AppData\Roaming\Microsoft\AddIns"
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Sub CopyFileForce(ByVal src As String, ByVal dst As String)
    ' Overwrite-compatible copy (handles existing file).
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.CopyFile src, dst, True
    If Err.Number <> 0 Then
        Err.Raise Err.Number, , "Failed to copy add-in to '" & dst & "': " & Err.Description
    End If
    On Error GoTo 0
End Sub

Private Function FindAddInByExactPath(ByVal fullPath As String) As AddIn
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If StrComp(ai.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindAddInByExactPath = ai
            Exit Function
        End If
    Next ai
End Function

Private Sub DisableStaleAddInRegistrations(ByVal addinBaseName As String, ByVal keepFullPath As String)
    ' Turn off any registrations for same name that DON'T match the canonical path
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If StrComp(ai.Name, addinBaseName, vbTextCompare) = 0 Then
            If StrComp(ai.FullName, keepFullPath, vbTextCompare) <> 0 Then
                On Error Resume Next
                ai.Installed = False
                Application.AddIns(ai.Name).Installed = False
                On Error GoTo 0
            End If
        End If
    Next ai
End Sub

