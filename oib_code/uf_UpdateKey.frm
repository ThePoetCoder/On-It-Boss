VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_UpdateKey 
   Caption         =   "Update API Key"
   ClientHeight    =   276
   ClientLeft      =   -72
   ClientTop       =   -348
   ClientWidth     =   2220
   OleObjectBlob   =   "uf_UpdateKey.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_UpdateKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' Center the form and load the stored API key into the textbox.
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Height = 64
        .Width = 352
        .txt_Key.PasswordChar = "*"
        .txt_Key.Value = mod_Env.ReadApiKeyFromEnv()
    End With
End Sub

Private Sub btn_Hide_Click()
    ' Toggle showing/hiding the API key characters.
    If Me.txt_Key.PasswordChar = "*" Then
        Me.txt_Key.PasswordChar = ""
    Else
        Me.txt_Key.PasswordChar = "*"
    End If
End Sub

Private Sub btn_Clear_Click()
    ' Clear the API key textbox.
    Me.txt_Key.Value = ""
End Sub

Private Sub btn_Update_Click()
    ' Validate and persist the API key to the .env file.
    Dim newKey As String
    newKey = Trim$(Me.txt_Key.Value)
    
    If Len(newKey) = 0 Then
        MsgBox "Please enter a valid API key.", vbExclamation
        Exit Sub
    End If
    
    mod_Env.WriteApiKeyToEnv newKey
    MsgBox "API key updated successfully!", vbInformation
    
    Unload Me
End Sub

