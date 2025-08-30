Attribute VB_Name = "mod_Config"
Option Explicit
Option Private Module

' One place for shared constants, endpoints, feature flags, and enums used across modules.
Public Const title As String = "On It, Boss! • "
Public Const OPENAI_MODELS_ENDPOINT As String = "https://api.openai.com/v1/models"
Public Const OPENAI_CHAT_ENDPOINT As String = "https://api.openai.com/v1/chat/completions"
Public Const DEBUG_CHAT As Boolean = False
Public Const HISTORY_SHEET_NAME As String = "_OIBHistory"

' History column indices for the _OIBHistory sheet.
Public Enum HistCols
    hcID = 1           ' Long (monotonic)
    hcTableName = 2    ' String (ListObject.Name)
    hcQueryName = 3    ' String (Sanitized query name for that table)
    hcTitle = 4        ' String (LLM "title")
    hcLanguage = 5     ' String (LLM "language", expect "m")
    hcCode = 6         ' String (Power Query M)
    hcCreatedAt = 7    ' Date (Now)
End Enum

