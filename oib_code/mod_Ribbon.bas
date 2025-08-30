Attribute VB_Name = "mod_Ribbon"
Option Explicit

Public Sub Ribbon_AskAI( _
    control As IRibbonControl _
):
    ' Show the main Ask AI panel from the ribbon callback.
    uf_Panel.Show
End Sub
Public Sub Ribbon_UpdateAPI( _
    control As IRibbonControl _
):
    ' Show the API key update form from the ribbon callback.
    uf_UpdateKey.Show
End Sub

