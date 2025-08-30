VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Panel 
   Caption         =   "On It, Boss!"
   ClientHeight    =   5088
   ClientLeft      =   60
   ClientTop       =   156
   ClientWidth     =   4488
   OleObjectBlob   =   "uf_Panel.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTable As ListObject
Private mDidInitTableContext As Boolean
Private Const FORM_OFFSET_PIXELS As Long = 15
Private Const BUTTON_FACE_BG As Long = vbButtonFace
Private Const RED_BG As Long = vbRed

Private Sub UserForm_Initialize()
    ' Initialize sizes, model combo placeholders, table context, and initial UI state.
    SetSizesAndPositions
    InitModelCombo
    InitTableContext_Once
    RefreshPanelState  ' sets caption, list, Ask button based on current table
End Sub

Private Sub UserForm_Activate()
    ' Refresh UI for current selection and lazily load real models.
    RefreshPanelState          ' keep UI in sync with current selection
    Me.txt_Chat.SetFocus
End Sub

Private Sub InitTableContext_Once()
    ' Build initial context; do the heavier “ensure” work only once.
    Dim lo As ListObject
    Set lo = mod_ExcelHelpers.ActiveListObjectOrNothing()
    If lo Is Nothing Then Exit Sub

    Set mTable = lo
    mod_History.History_EnsureSeedForTable mTable.Parent.Parent, mTable
    mod_SchemaQuery.EnsureQueryAndSheetForTable mTable

    ' If user opened panel on source-only sheet, jump them to the output once (same behavior).
    Dim qName As String: qName = mod_String.SanitizeName(mTable.name)
    If Not mod_SchemaQuery.IsListObjectBoundToQuery(mTable, qName) Then
        mod_SchemaQuery.EnsureAndActivateQueryOutputForTable mTable
        ' After jumping, re-evaluate the active table handle:
        Set mTable = mod_ExcelHelpers.ActiveListObjectOrNothing()
    End If

    mDidInitTableContext = True
End Sub

Private Sub RefreshPanelState()
    ' Sync the panel controls with the currently active table (or no table).
    Dim lo As ListObject
    Set lo = mod_ExcelHelpers.ActiveListObjectOrNothing()

    If lo Is Nothing Then
        Set mTable = Nothing
        UI_NoTable
        Exit Sub
    End If

    ' If selection switched tables, update our context
    If (mTable Is Nothing) Or (lo.name <> mTable.name) Then
        Set mTable = lo
        ' Only the very first time we do “ensure query + jump”. On later Activates, just refresh UI.
        If Not mDidInitTableContext Then InitTableContext_Once
        UI_ForTable mTable
    Else
        ' Same table; still make sure UI is enabled and captioned right
        UI_ForTable mTable
    End If
End Sub

Private Sub UI_NoTable()
    ' Configure UI for "no table selected" state.
    Me.Caption = mod_Config.title & "(no table selected)"
    Me.btn_AskAI.Enabled = False
    Me.lst_Prev.clear
    
    Static warned As Boolean
    If Not warned Then
        MsgBox "No Excel table detected on this sheet. Insert a table (Ctrl+T) or select a sheet that has one.", vbInformation
        warned = True
    End If
End Sub

Private Sub UI_ForTable( _
    ByVal lo As ListObject _
)
    ' Configure UI for a specific ListObject and populate history list.
    Me.Caption = mod_Config.title & lo.name
    Me.btn_AskAI.Enabled = True
    mod_History.Panel_FillListForTable Me, lo
End Sub

Sub InitModelCombo()
    ' Seed the model ComboBox with models.
    With Me.cmb_Model
        .clear
        .AddItem "gpt-5-nano"
        .AddItem "gpt-5-mini"
        .AddItem "gpt-5"
        .Value = "gpt-5-nano"
    End With
End Sub

Private Sub SetSizesAndPositions()
    ' Set fixed positions and sizes of panel controls.
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + FORM_OFFSET_PIXELS
        .Top = Application.Top + FORM_OFFSET_PIXELS
        .Width = 288
        .Height = 333

        .lst_Prev.Top = 6
        .lst_Prev.Left = 6
        .lst_Prev.Height = 204.85
        .lst_Prev.Width = 263.25

        .txt_Chat.Top = 216
        .txt_Chat.Left = 6
        .txt_Chat.Height = 54
        .txt_Chat.Width = 216

        .btn_AskAI.Top = 216
        .btn_AskAI.Left = 223.05
        .btn_AskAI.Height = 54
        .btn_AskAI.Width = 46.2

        .lbl_Model.Top = 282
        .lbl_Model.Left = 6
        .lbl_Model.Height = 18
        .lbl_Model.Width = 48

        .cmb_Model.Top = 276
        .cmb_Model.Left = 58.15
        .cmb_Model.Height = 24
        .cmb_Model.Width = 211.85
    End With
End Sub


Private Sub btn_AskAI_Click()
    ' Send the user’s request to the model and apply the returned M to the active table.
    Dim txt As String: txt = Trim$(Me.txt_Chat.text)
    If Len(txt) = 0 Then
        MsgBox "Type a request.", vbExclamation
        Exit Sub
    End If
    
    Me.btn_AskAI.BackColor = RED_BG
    Me.btn_AskAI.Caption = "Asking"
    mod_Core.ChatOIB txt, mTable

    ' Refresh list to show the new entry at the bottom (newest)
    mod_History.Panel_FillListForTable Me, mTable
    ' clear input
    Me.txt_Chat.text = ""
    Me.btn_AskAI.BackColor = BUTTON_FACE_BG
    Me.btn_AskAI.Caption = "Ask AI"
    
    If mTable Is Nothing Then
        MsgBox "No table selected. Please select a sheet with a table and try again.", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub lst_Prev_Click()
    ' Load and immediately apply the selected history item’s M code.
    If mTable Is Nothing Then Exit Sub
    If Me.lst_Prev.ListIndex < 0 Then Exit Sub

    Dim id As Long
    id = CLng(Me.lst_Prev.List(Me.lst_Prev.ListIndex, 0)) ' hidden ID col

    Dim title As String, lang As String
    Dim mcode As String
    mcode = mod_History.History_GetCodeById(mTable.Parent.Parent, id, title, lang)
    If Len(mcode) = 0 Then
        MsgBox "Couldn't load M for the selected item.", vbExclamation
        Exit Sub
    End If
    If LCase$(lang) <> "m" Then
        MsgBox "Selected entry isn't language='m'.", vbExclamation
        Exit Sub
    End If

    ' Immediately apply it (your requirement)
    mod_SchemaQuery.ApplyMToTable mTable, mcode, title
End Sub

Private Sub lst_Prev_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
)
    ' Handle Delete key to remove a selected history item from the list.
    If KeyCode = vbKeyDelete Then
        
        If mTable Is Nothing Then Exit Sub
        If Me.lst_Prev.ListIndex < 0 Then
            MsgBox "Select an item to delete.", vbExclamation
            Exit Sub
        End If
    
        Dim id As Long
        id = CLng(Me.lst_Prev.List(Me.lst_Prev.ListIndex, 0))
    
        If MsgBox("Delete the selected history item?", vbQuestion + vbYesNo, "Delete") = vbYes Then
            mod_History.History_DeleteById mTable.Parent.Parent, id
            mod_History.Panel_FillListForTable Me, mTable
        End If
    End If
End Sub

Private Sub txt_Chat_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' If Enter is pressed without Shift, insert a newline instead of moving focus
    If KeyCode = vbKeyReturn Then
        Dim selStart As Long
        selStart = Me.txt_Chat.selStart
        Me.txt_Chat.text = Left$(Me.txt_Chat.text, selStart) & vbCrLf & Mid$(Me.txt_Chat.text, selStart + 1)
        Me.txt_Chat.selStart = selStart + 2
        KeyCode = 0
    End If
End Sub

Public Property Get SelectedModel() As String
    SelectedModel = Me.cmb_Model.Value
End Property

Public Property Get SelectedTemperature() As Double
    SelectedTemperature = 1#
End Property


