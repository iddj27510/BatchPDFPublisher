Attribute VB_Name = "modState"
Option Explicit
' Module    : modState
' Purpose   : Global runtime state, registry access, reset
'             orchestration, config validation, registry schema
'             migration.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' All Public Const declarations live in modConstants. This module
' owns only mutable runtime state (CLR_*, g_*), registry cache,
' reset orchestration, config validation, and registry schema
' migration.
'
' Rule: CLR_* are Public variables, not Const.
' Rule: all constants live in modConstants.
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in procedure signatures and multi-line string
'   concatenations). No functional changes.
'
' v9.19.6 (STATE_ValidateConfig Function -> Sub):
'   The function always returned True (historical — original v9.18.5
'   implementation never set bOK = False; no caller used the return
'   value). Converted to Sub to signal the API contract accurately:
'   the validation result is carried by the ByRef sReport parameter.
'   All callers updated.
'
' v9.19.3 (Header standardization):
'   Header converted from Style A to Style B. English throughout.
'
' v9.19.3 (STATE_ValidateConfig refactor):
'   Refactored from a monolithic 6-block function (CC ~12) to a
'   flat orchestrator (CC ~1) plus four small typed validators:
'     STATE_ValidateEnum, STATE_ValidateBool,
'     STATE_ValidateIntRange, STATE_ValidateLang.
'
' v9.19.3 (STATE_BackupPreferences unified writer):
'   STATE_BackupPreferences now uses modSandbox.PRJ_WriteTextFile
'   and PRJ_PathCombine. Same canonical writer as modError and
'   modReportExport.
'
' Requires  : PRJ_FSO (modCOM), PRJ_SandboxPath, PRJ_PathCombine,
'             PRJ_WriteTextFile (modSandbox),
'             PRJ_LogError, PRJ_LogInfo, PRJ_LogWarn (modLog),
'             L, LANG_SupportedCodes (modLang),
'             REG_KEY_* and STATE_SCHEMA_VERSION (modConstants).
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-line string
'             concatenations are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public CLR_BG           As Long
Public CLR_SURFACE      As Long
Public CLR_SEP          As Long
Public CLR_TEXT_PRI     As Long
Public CLR_TEXT_SEC     As Long
Public CLR_ACCENT       As Long
Public CLR_ACCENT_TXT   As Long
Public CLR_BTN_SEC      As Long
Public CLR_BTN_SEC_TXT  As Long
Public CLR_ERROR        As Long
Public CLR_OK           As Long
Public CLR_WARN         As Long

Public CLR_RUN_BG       As Long
Public CLR_RUN_TXT      As Long
Public CLR_RUN_DIS      As Long
Public CLR_RUN_DIS_TXT  As Long

Public BTN_HOVER_RUN    As Long
Public BTN_PRESS_RUN    As Long
Public BTN_HOVER_SEC    As Long
Public BTN_PRESS_SEC    As Long
Public BTN_HOVER_ACCENT As Long
Public BTN_PRESS_ACCENT As Long

Public CLR_TEMP_OK      As Long
Public CLR_TEMP_WARN    As Long
Public CLR_TEMP_INFO    As Long

Public CLR_PROCESS      As Long

Public g_sLastFolder     As String
Public g_bSubfolders     As Boolean
Public g_bForceOverwrite As Boolean

Public g_sTheme          As String
Public g_bThemeApplied   As Boolean
Public g_bCurrentDark    As Boolean
Public g_bHighContrast   As Boolean

Public g_colFiles        As Object
Public g_colLog          As Object
Public g_nTotal          As Long
Public g_nDone           As Long
Public g_bCancel         As Boolean
Public g_bInProcessLoop  As Boolean
Public g_bRetryMode      As Boolean
Public g_bIsProcessing   As Boolean
Public g_dTimeStart      As Double
Public g_dTimeEnd        As Double

Public g_oProgress       As Object
Public g_oFSO            As Object
Public g_oApprentice     As Object

Public g_dPageCache      As Object
Public g_dSheetCache     As Object

Public g_colHoverBtns    As Object
Public g_bHoverRunEnabled As Boolean

Public g_sActiveFilter   As String

Public g_nLogLevel       As Long

Public g_bAuditMode      As Boolean
Public g_bAnalyzeIDW     As Boolean
Public g_bAnalyzePDF     As Boolean

Public g_sCallStack      As String
Public g_sAbortReason    As String

Public g_sCurrentRelPath As String

Public g_nTempCounter    As Long

Private m_oRegCache As Object

Public Function REG_Get(ByVal sKey As String, _
                         Optional ByVal sDefault As String = "") As String
    On Error Resume Next

    If m_oRegCache Is Nothing Then
        Err.Clear
        Set m_oRegCache = CreateObject("Scripting.Dictionary")
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            REG_Get = sDefault
            Exit Function
        End If
        m_oRegCache.CompareMode = 1
    End If

    If m_oRegCache.Exists(sKey) Then
        REG_Get = CStr(m_oRegCache(sKey))
        Exit Function
    End If

    Dim sVal As String
    sVal = GetSetting(REG_SECTION, REG_PREFS, sKey, sDefault)
    m_oRegCache(sKey) = sVal
    REG_Get = sVal
    On Error GoTo 0
End Function

Public Sub REG_Set(ByVal sKey As String, ByVal sValue As String)
    On Error Resume Next

    If m_oRegCache Is Nothing Then
        Err.Clear
        Set m_oRegCache = CreateObject("Scripting.Dictionary")
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        m_oRegCache.CompareMode = 1
    End If

    SaveSetting REG_SECTION, REG_PREFS, sKey, sValue
    m_oRegCache(sKey) = sValue
    On Error GoTo 0
End Sub

Public Sub REG_ClearCache()
    On Error Resume Next
    If Not m_oRegCache Is Nothing Then m_oRegCache.RemoveAll
    On Error GoTo 0
End Sub

Public Sub STATE_Init()
    REG_ClearCache

    g_sLastFolder = ""
    g_bSubfolders = True
    g_bForceOverwrite = False

    g_sTheme = "I"
    g_bThemeApplied = False
    g_bCurrentDark = False
    g_bHighContrast = False

    Set g_colFiles = Nothing
    Set g_colLog = Nothing
    g_nTotal = 0
    g_nDone = 0
    g_bCancel = False
    g_bInProcessLoop = False
    g_bRetryMode = False
    g_bIsProcessing = False
    g_dTimeStart = 0
    g_dTimeEnd = 0

    Set g_oProgress = Nothing
    Set g_oApprentice = Nothing

    Set g_dPageCache = Nothing
    Set g_dSheetCache = Nothing

    Set g_colHoverBtns = Nothing
    g_bHoverRunEnabled = False

    g_sActiveFilter = "All"
    g_nLogLevel = LOG_INFO

    g_bAuditMode = False
    g_bAnalyzeIDW = False
    g_bAnalyzePDF = False

    g_sCallStack = ""
    g_sAbortReason = ""
    g_sCurrentRelPath = ""
    g_nTempCounter = 0

    STATE_MigrateSchema
End Sub

Public Sub STATE_Reset()
    STATE_Init
End Sub

Public Sub STATE_ResetPreferences()
    STATE_BackupPreferences

    g_bSubfolders = True
    g_bForceOverwrite = False
    g_sLastFolder = ""

    REG_Set REG_KEY_PATH, ""
    REG_Set REG_KEY_SUB, "True"
    REG_Set REG_KEY_FORCE, "False"

    PRJ_LogInfo "AUDIT", "STATE_ResetPreferences: user reset all preferences"
End Sub

' STATE_ValidateConfig   validate every registry key against its
'   expected value set. Invalid values are reset to defaults and
'   logged as WARN. The result is written to sReport (ByRef).
'
'   v9.19.6: converted from Function (always returned True) to Sub.
'   No caller used the return value; the ByRef sReport parameter
'   carries the validation result.
'
'   The orchestrator is a table-driven list of four typed
'   validators. Adding a new registry key means adding one line
'   here, not editing a monolithic block.
Public Sub STATE_ValidateConfig(ByRef sReport As String)
    Dim sOut As String
    sOut = "Config validation:" & vbCrLf

    STATE_ValidateEnum REG_KEY_THEME, "I", "I|S|L|D", "Theme", sOut
    STATE_ValidateBool REG_KEY_SUB, "True", "Sub", sOut
    STATE_ValidateBool REG_KEY_FORCE, "False", "Force", sOut
    STATE_ValidateIntRange REG_KEY_LOGLEVEL, "1", _
                           LOG_DEBUG, LOG_CRITICAL, "LogLevel", sOut
    STATE_ValidateBool REG_KEY_LOGJSON, "False", "LogJson", sOut
    STATE_ValidateLang REG_KEY_LANG, "", "Lang", sOut

    sReport = sOut
End Sub

' STATE_ValidateEnum   validate a registry value against a
'   delimited allow-list. sAllowed is a "|"-separated string of
'   accepted values; sVal must match exactly one entry.
'
'   Empty sVal always fails. Callers that permit an empty value
'   (e.g. Lang, which accepts "" for auto-detect) must not use
'   this validator — use STATE_ValidateLang instead.
'
'   v9.19.6 fix: prior implementation used InStr(sAllowed, sVal),
'   which accepted substrings — "IS" and "ISL" both passed
'   against sAllowed="ISLD". Now uses delimited exact match.
Private Sub STATE_ValidateEnum(ByVal sKey As String, _
                                ByVal sDefault As String, _
                                ByVal sAllowed As String, _
                                ByVal sLabel As String, _
                                ByRef sOut As String)
    Dim sVal As String
    sVal = REG_Get(sKey, sDefault)

    If Len(sVal) > 0 Then
        If InStr("|" & sAllowed & "|", "|" & sVal & "|") > 0 Then
            sOut = sOut & "  [PASS] " & sLabel & " = " & sVal & vbCrLf
            Exit Sub
        End If
    End If

    sOut = sOut & "  [WARN] " & sLabel & " invalid (" & sVal & _
           "), reset to " & sDefault & vbCrLf
    REG_Set sKey, sDefault
    PRJ_LogWarn "STATE_ValidateConfig", _
        "Invalid " & sLabel & " '" & sVal & "' reset"
End Sub

' STATE_ValidateBool   validate a registry value as "true" or "false"
'   (case-insensitive). Anything else resets to default.
Private Sub STATE_ValidateBool(ByVal sKey As String, _
                                ByVal sDefault As String, _
                                ByVal sLabel As String, _
                                ByRef sOut As String)
    Dim sVal As String
    sVal = LCase$(REG_Get(sKey, sDefault))

    If sVal = "true" Or sVal = "false" Then
        sOut = sOut & "  [PASS] " & sLabel & " = " & sVal & vbCrLf
        Exit Sub
    End If

    sOut = sOut & "  [WARN] " & sLabel & " invalid (" & sVal & _
           "), reset to " & sDefault & vbCrLf
    REG_Set sKey, sDefault
    PRJ_LogWarn "STATE_ValidateConfig", _
        "Invalid " & sLabel & " '" & sVal & "' reset"
End Sub

' STATE_ValidateIntRange   validate a registry value as an integer in
'   [nMin, nMax]. Non-numeric values and out-of-range values both
'   reset to default, with distinct WARN messages.
Private Sub STATE_ValidateIntRange(ByVal sKey As String, _
                                    ByVal sDefault As String, _
                                    ByVal nMin As Long, _
                                    ByVal nMax As Long, _
                                    ByVal sLabel As String, _
                                    ByRef sOut As String)
    Dim sVal As String
    sVal = REG_Get(sKey, sDefault)

    If Not IsNumeric(sVal) Then
        sOut = sOut & "  [WARN] " & sLabel & " non-numeric (" & sVal & _
               "), reset to " & sDefault & vbCrLf
        REG_Set sKey, sDefault
        PRJ_LogWarn "STATE_ValidateConfig", _
            "Non-numeric " & sLabel & " '" & sVal & "' reset"
        Exit Sub
    End If

    Dim nVal As Long
    nVal = CLng(sVal)

    If nVal < nMin Or nVal > nMax Then
        sOut = sOut & "  [WARN] " & sLabel & " out of range (" & sVal & _
               "), reset to " & sDefault & vbCrLf
        REG_Set sKey, sDefault
        PRJ_LogWarn "STATE_ValidateConfig", _
            sLabel & " out of range '" & sVal & "' reset"
        Exit Sub
    End If

    sOut = sOut & "  [PASS] " & sLabel & " = " & sVal & vbCrLf
End Sub

' STATE_ValidateLang   validate a registry value as either empty
'   (auto-detect) or a code present in LANG_SupportedCodes.
'   Case-insensitive match. Invalid codes reset to empty.
Private Sub STATE_ValidateLang(ByVal sKey As String, _
                                ByVal sDefault As String, _
                                ByVal sLabel As String, _
                                ByRef sOut As String)
    Dim sVal As String
    sVal = LCase$(REG_Get(sKey, sDefault))

    If Len(sVal) = 0 Then
        sOut = sOut & "  [PASS] " & sLabel & " = (auto)" & vbCrLf
        Exit Sub
    End If

    Dim vCodes As Variant
    Dim i As Long
    vCodes = LANG_SupportedCodes()

    For i = LBound(vCodes) To UBound(vCodes)
        If sVal = CStr(vCodes(i)) Then
            sOut = sOut & "  [PASS] " & sLabel & " = " & sVal & vbCrLf
            Exit Sub
        End If
    Next i

    sOut = sOut & "  [WARN] " & sLabel & " invalid (" & sVal & _
           "), reset to auto" & vbCrLf
    REG_Set sKey, sDefault
    PRJ_LogWarn "STATE_ValidateConfig", _
        "Invalid " & sLabel & " '" & sVal & "' reset"
End Sub

' STATE_MigrateSchema   read the persisted registry schema version
'   and bring it up to date. Called once per session from STATE_Init.
Public Sub STATE_MigrateSchema()
    On Error GoTo EH

    Dim sPersisted As String
    Dim nPersisted As Long

    sPersisted = REG_Get(REG_KEY_SCHEMAVERSION, "")
    If Len(sPersisted) = 0 Then
        REG_Set REG_KEY_SCHEMAVERSION, CStr(STATE_SCHEMA_VERSION)
        PRJ_LogInfo "STATE_MigrateSchema", _
            "Schema version initialised to " & STATE_SCHEMA_VERSION
        Exit Sub
    End If

    If Not IsNumeric(sPersisted) Then
        PRJ_LogWarn "STATE_MigrateSchema", _
            "Non-numeric schema version '" & sPersisted & _
            "', resetting to " & STATE_SCHEMA_VERSION
        REG_Set REG_KEY_SCHEMAVERSION, CStr(STATE_SCHEMA_VERSION)
        Exit Sub
    End If

    nPersisted = CLng(sPersisted)

    If nPersisted = STATE_SCHEMA_VERSION Then Exit Sub

    If nPersisted > STATE_SCHEMA_VERSION Then
        PRJ_LogWarn "STATE_MigrateSchema", _
            "Registry schema v" & nPersisted & _
            " is newer than code schema v" & STATE_SCHEMA_VERSION & _
            "; no migration attempted"
        Exit Sub
    End If

    Select Case nPersisted
        Case 0
            ' Reserved. Never used.
        Case 1
            ' v1 -> v2: added LogJson key with default False.
            PRJ_LogInfo "STATE_MigrateSchema", _
                "v1 -> v2: LogJson key default False"
        Case Else
            PRJ_LogWarn "STATE_MigrateSchema", _
                "No migration path from schema v" & nPersisted & _
                " to v" & STATE_SCHEMA_VERSION
    End Select

    REG_Set REG_KEY_SCHEMAVERSION, CStr(STATE_SCHEMA_VERSION)
    PRJ_LogInfo "AUDIT", _
        "Registry schema migrated from v" & nPersisted & _
        " to v" & STATE_SCHEMA_VERSION
    Exit Sub

EH:
    PRJ_LogWarn "STATE_MigrateSchema", "Migration error: " & Err.Description
End Sub

Public Sub STATE_BackupPreferences()
    Dim sPath As String
    Dim sBody As String

    On Error GoTo EH

    PRJ_EnsureSandbox

    sPath = PRJ_PathCombine(PRJ_SandboxPath(), _
            "prefs_backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt")

    sBody = "Batch PDF Publisher - Preferences Backup" & vbCrLf
    sBody = sBody & "Timestamp: " & Format$(Now, "yyyy-mm-dd HH:mm:ss") & vbCrLf
    sBody = sBody & "Version:   " & PRJ_VERSION & vbCrLf
    sBody = sBody & "SchemaVer: " & REG_Get(REG_KEY_SCHEMAVERSION, "") & vbCrLf
    sBody = sBody & vbCrLf

    sBody = sBody & "Theme       = " & REG_Get(REG_KEY_THEME, "I") & vbCrLf
    sBody = sBody & "Path        = " & REG_Get(REG_KEY_PATH, "") & vbCrLf
    sBody = sBody & "Sub         = " & REG_Get(REG_KEY_SUB, "True") & vbCrLf
    sBody = sBody & "Force       = " & REG_Get(REG_KEY_FORCE, "False") & vbCrLf
    sBody = sBody & "Audit       = " & REG_Get(REG_KEY_AUDIT, "False") & vbCrLf
    sBody = sBody & "LogLevel    = " & REG_Get(REG_KEY_LOGLEVEL, "1") & vbCrLf
    sBody = sBody & "LogJson     = " & REG_Get(REG_KEY_LOGJSON, "False") & vbCrLf
    sBody = sBody & "Lang        = " & REG_Get(REG_KEY_LANG, "") & vbCrLf

    If Not PRJ_WriteTextFile(sPath, sBody) Then
        PRJ_LogWarn "STATE_BackupPreferences", "Backup write failed: " & sPath
        Exit Sub
    End If

    PRJ_LogInfo "AUDIT", "Preferences backed up: " & sPath
    Exit Sub

EH:
    PRJ_LogWarn "STATE_BackupPreferences", "Backup failed: " & Err.Description
End Sub

Public Function STATE_AssertGlobals(ByRef sReport As String) As Boolean
    Dim bOK As Boolean
    Dim sOut As String

    bOK = True
    sOut = "Global state assertion:" & vbCrLf

    If g_colFiles Is Nothing Then
        sOut = sOut & "  [WARN] g_colFiles is Nothing (batch not initialized)" & vbCrLf
    Else
        sOut = sOut & "  [PASS] g_colFiles count = " & g_colFiles.Count & vbCrLf
    End If

    If g_oFSO Is Nothing Then
        sOut = sOut & "  [WARN] g_oFSO is Nothing (lazy, will create on demand)" & vbCrLf
    Else
        sOut = sOut & "  [PASS] g_oFSO created" & vbCrLf
    End If

    If Len(g_sTheme) = 0 Then
        sOut = sOut & "  [FAIL] g_sTheme is empty" & vbCrLf
        bOK = False
    ElseIf InStr("ISLD", g_sTheme) = 0 Then
        sOut = sOut & "  [FAIL] g_sTheme invalid: " & g_sTheme & vbCrLf
        bOK = False
    Else
        sOut = sOut & "  [PASS] g_sTheme = " & g_sTheme & vbCrLf
    End If

    If g_nLogLevel < LOG_DEBUG Or g_nLogLevel > LOG_CRITICAL Then
        sOut = sOut & "  [FAIL] g_nLogLevel out of range: " & g_nLogLevel & vbCrLf
        bOK = False
    Else
        sOut = sOut & "  [PASS] g_nLogLevel = " & g_nLogLevel & vbCrLf
    End If

    If Len(g_sActiveFilter) = 0 Then
        sOut = sOut & "  [FAIL] g_sActiveFilter is empty" & vbCrLf
        bOK = False
    ElseIf InStr("AllErrSkipOkNoIdwNoDWG", g_sActiveFilter) = 0 Then
        sOut = sOut & "  [WARN] g_sActiveFilter unknown value: " & g_sActiveFilter & vbCrLf
    Else
        sOut = sOut & "  [PASS] g_sActiveFilter = " & g_sActiveFilter & vbCrLf
    End If

    If g_nTotal < 0 Then
        sOut = sOut & "  [FAIL] g_nTotal negative: " & g_nTotal & vbCrLf
        bOK = False
    Else
        sOut = sOut & "  [PASS] g_nTotal = " & g_nTotal & vbCrLf
    End If

    If g_nDone < 0 Then
        sOut = sOut & "  [FAIL] g_nDone negative: " & g_nDone & vbCrLf
        bOK = False
    ElseIf g_nTotal > 0 And g_nDone > g_nTotal Then
        sOut = sOut & "  [FAIL] g_nDone (" & g_nDone & ") > g_nTotal (" & g_nTotal & ")" & vbCrLf
        bOK = False
    Else
        sOut = sOut & "  [PASS] g_nDone = " & g_nDone & vbCrLf
    End If

    If g_dTimeStart > 0 And g_dTimeEnd > 0 And g_dTimeEnd < g_dTimeStart Then
        sOut = sOut & "  [FAIL] g_dTimeEnd (" & g_dTimeEnd & ") < g_dTimeStart (" & g_dTimeStart & ")" & vbCrLf
        bOK = False
    Else
        sOut = sOut & "  [PASS] time range OK" & vbCrLf
    End If

    If g_nTempCounter < 0 Then
        sOut = sOut & "  [FAIL] g_nTempCounter negative: " & g_nTempCounter & vbCrLf
        bOK = False
    Else
        sOut = sOut & "  [PASS] g_nTempCounter = " & g_nTempCounter & vbCrLf
    End If

    sReport = sOut
    STATE_AssertGlobals = bOK
End Function

Public Sub ResetAllSettings()
    Dim oFSO     As Object
    Dim sSandbox As String
    Dim sMsg     As String

    On Error GoTo EH

    STATE_BackupPreferences

    DeleteSetting REG_SECTION, REG_PREFS
    REG_ClearCache

    Set oFSO = PRJ_FSO()
    sSandbox = PRJ_SandboxPath()
    If Not oFSO Is Nothing Then
        If oFSO.FolderExists(sSandbox) Then
            oFSO.DeleteFolder sSandbox, True
        End If
    End If

    STATE_Init

    sMsg = L("reset.msg.all_defaults") & vbCrLf & vbCrLf
    sMsg = sMsg & L("reset.lbl.registry") & ": " & L("reset.state.deleted") & vbCrLf
    sMsg = sMsg & L("reset.lbl.sandbox") & ": " & sSandbox
    sMsg = sMsg & " - " & L("reset.state.deleted") & vbCrLf
    sMsg = sMsg & L("reset.lbl.state") & ": " & L("reset.state.reinitialized")

    MsgBox sMsg, vbInformation, L("title")
    Exit Sub

EH:
    PRJ_LogError "ResetAllSettings", Err.Description
    MsgBox L("reset.msg.failed") & ": " & Err.Description, vbExclamation, L("title")
End Sub

Public Sub ResetOnlyRegistry()
    On Error GoTo EH
    STATE_BackupPreferences
    DeleteSetting REG_SECTION, REG_PREFS
    REG_ClearCache
    STATE_Reset
    STATE_Init
    MsgBox L("reset.msg.registry_complete"), vbInformation, L("title")
    Exit Sub
EH:
    PRJ_LogError "ResetOnlyRegistry", Err.Description
    MsgBox L("reset.msg.failed") & ": " & Err.Description, vbExclamation, L("title")
End Sub

Public Sub ResetOnlySandbox()
    Dim oFSO     As Object
    Dim sSandbox As String

    On Error GoTo EH
    Set oFSO = PRJ_FSO()
    sSandbox = PRJ_SandboxPath()
    If Not oFSO Is Nothing Then
        If oFSO.FolderExists(sSandbox) Then
            oFSO.DeleteFolder sSandbox, True
        End If
    End If
    MsgBox L("reset.msg.sandbox_complete") & ": " & sSandbox, vbInformation, L("title")
    Exit Sub
EH:
    PRJ_LogError "ResetOnlySandbox", Err.Description
    MsgBox L("reset.msg.failed") & ": " & Err.Description, vbExclamation, L("title")
End Sub

' END: modState v9.19.9

