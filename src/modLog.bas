Attribute VB_Name = "modLog"
Option Explicit
' Module    : modLog
' Purpose   : Structured file logging with level filtering, JSON mode,
'             audit sub-log, and multi-archive rotation.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in procedure signatures, multi-line string
'   concatenations and multi-line Debug.Print statements).
'   No functional changes.
'
' v9.19.6: Format -> Format$ in LOG_FormatLine / LOG_Close
'          (project §3.19 convention — avoid Variant wrap on hot path).
'          Header sync to PRJ_VERSION.
' v9.19.3 (§4.3 Silent failures):
'   Silent Err.Clear in LOG_WriteAuditRecord and LOG_OpenInternal
'   now emits Debug.Print instead of swallowing the error. Using
'   Debug.Print — not PRJ_LogDebug — because calling back into
'   LOG_Write from inside LOG_WriteAuditRecord would re-enter
'   LOG_Rotate when the buffer threshold is hit, and the audit
'   write path is itself the fallback for a failing main log.
'
' v9.18.5 (§16.2 Log levels):
'   Added LOG_CRITICAL. PRJ_Fatal now emits at CRITICAL so fatal
'   errors can be filtered separately from ordinary errors.
' v9.18.5 (§16.4 Rotation):
'   Multi-archive rotation. Keeps LOG_ROTATE_MAX archives
'   (BatchPDF_Log_prev_1, prev_2, ...). Prior code kept only one.
' v9.18.5 (§16.6 Audit trail):
'   AUDIT context now also writes to a separate audit file
'   (BatchPDF_Audit_<PID>.txt) with 90-day retention, distinct from
'   the 24h main log retention.
' v9.18.5 (§16.8 Structured logging):
'   Optional JSONL mode (one JSON object per line) via LogJson
'   registry key. Off by default to preserve existing log parsers.
' v9.18.5 (§16.10 Runtime reload):
'   Added LOG_ReloadLevel — re-reads LogLevel from registry without
'   restart. Wired to PRJ_ReloadLogLevel in modEntry.
' v9.18.0 (§5.17 Log sanitization):
'   LOG_SanitizePII — masks %USERPROFILE% and \Users\<name>\.
' v9.17.0: Removed LOG_GetPath — redundant alias for PRJ_LogPath.
' v9.14.8 (4, H.18): LOG_Asciify filter - non-ASCII -> ?.
'
' Reentrancy Guards (this module):
'   m_bInOpen    — breaks cycle 1 (LOG_Open -> ... -> LOG_Open)
'   m_bInRotate  — breaks cycle 4 (LOG_Write -> ... -> LOG_Rotate)
'   PRJ_LogSubsystemBusy() — public flag used by modSandbox (cycle 2)
'
' Requires  : PRJ_SandboxPath, PRJ_PathCombine (modSandbox).
'             PRJ_FSO (modCOM).
'             g_nLogLevel, LOG_DEBUG/INFO/WARN/ERROR/CRITICAL,
'             LOG_MAX_BYTES, LOG_FLUSH_EVERY, LOG_FILE_PREFIX,
'             LOG_AUDIT_PREFIX, LOG_ROTATE_MAX, PRJ_VERSION
'             (modConstants / modState).
'             REG_Get (modState).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures, multi-line string concatenations
'             and multi-line Debug.Print statements are excepted
'             per Doc 01 / Doc 06.
'             No decorative separator comment lines.

#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private m_sLogPath     As String
Private m_bOpen        As Boolean
Private m_nFileNum     As Integer
Private m_bInRotate    As Boolean
Private m_bInOpen      As Boolean
Private m_bWriteFailed As Boolean
Private m_aLogBuf()    As String
Private m_nBufCount    As Long

Private m_sAuditPath     As String
Private m_bAuditOpen     As Boolean
Private m_nAuditFileNum  As Integer

Private m_bJsonMode    As Boolean
Private m_bJsonProbed  As Boolean

Public Function PRJ_LogPath() As String
    Dim sPid As String
    sPid = Format$(GetCurrentProcessId(), "00000")
    PRJ_LogPath = PRJ_PathCombine(PRJ_SandboxPath(), LOG_FILE_PREFIX & sPid & ".txt")
End Function

Public Function PRJ_LogAuditPath() As String
    Dim sPid As String
    sPid = Format$(GetCurrentProcessId(), "00000")
    PRJ_LogAuditPath = PRJ_PathCombine(PRJ_SandboxPath(), LOG_AUDIT_PREFIX & sPid & ".txt")
End Function

' LOG_GetArchivePath   return the archive path for a given rotation
'   slot. slot 1 is the most recent archive, slot LOG_ROTATE_MAX is
'   the oldest. Slot 0 is not valid — the live log uses PRJ_LogPath.
' Params  : nSlot — Long, 1 to LOG_ROTATE_MAX
' Returns : String — full path to archive slot
Private Function LOG_GetArchivePath(ByVal nSlot As Long) As String
    Dim sPid As String
    sPid = Format$(GetCurrentProcessId(), "00000")
    LOG_GetArchivePath = PRJ_PathCombine(PRJ_SandboxPath(), _
                          LOG_FILE_PREFIX & "prev" & CStr(nSlot) & "_" & sPid & ".txt")
End Function

Public Function PRJ_LogSubsystemBusy() As Boolean
    PRJ_LogSubsystemBusy = m_bInOpen Or m_bInRotate
End Function

Public Function PRJ_LogWriteFailed() As Boolean
    PRJ_LogWriteFailed = m_bWriteFailed
End Function

Private Function LOG_IsJsonMode() As Boolean
    If Not m_bJsonProbed Then
        On Error Resume Next
        Err.Clear
        m_bJsonMode = (LCase$(REG_Get(REG_KEY_LOGJSON, "False")) = "true")
        If Err.Number <> 0 Then
            Err.Clear
            m_bJsonMode = False
        End If
        On Error GoTo 0
        m_bJsonProbed = True
    End If
    LOG_IsJsonMode = m_bJsonMode
End Function

Private Function LOG_JsonEscape(ByVal s As String) As String
    Dim sOut As String
    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim nCode As Long

    sOut = ""
    n = Len(s)
    For i = 1 To n
        ch = Mid$(s, i, 1)
        nCode = AscW(ch)
        If nCode < 0 Then nCode = nCode + 65536
        Select Case nCode
            Case 34:   sOut = sOut & "\"""
            Case 92:   sOut = sOut & "\\"
            Case 8:    sOut = sOut & "\b"
            Case 9:    sOut = sOut & "\t"
            Case 10:   sOut = sOut & "\n"
            Case 12:   sOut = sOut & "\f"
            Case 13:   sOut = sOut & "\r"
            Case 0 To 31
                sOut = sOut & "\u" & Right$("0000" & Hex$(nCode), 4)
            Case Else: sOut = sOut & ch
        End Select
    Next i
    LOG_JsonEscape = sOut
End Function

' LOG_LevelName   convert numeric level to canonical uppercase name.
Private Function LOG_LevelName(ByVal nLevel As Long) As String
    Select Case nLevel
        Case LOG_DEBUG:    LOG_LevelName = "DEBUG"
        Case LOG_INFO:     LOG_LevelName = "INFO"
        Case LOG_WARN:     LOG_LevelName = "WARN"
        Case LOG_ERROR:    LOG_LevelName = "ERROR"
        Case LOG_CRITICAL: LOG_LevelName = "CRITICAL"
        Case Else:         LOG_LevelName = "INFO"
    End Select
End Function

Private Function LOG_LevelNamePadded(ByVal nLevel As Long) As String
    Select Case nLevel
        Case LOG_DEBUG:    LOG_LevelNamePadded = "DEBUG   "
        Case LOG_INFO:     LOG_LevelNamePadded = "INFO    "
        Case LOG_WARN:     LOG_LevelNamePadded = "WARN    "
        Case LOG_ERROR:    LOG_LevelNamePadded = "ERROR   "
        Case LOG_CRITICAL: LOG_LevelNamePadded = "CRITICAL"
        Case Else:         LOG_LevelNamePadded = "INFO    "
    End Select
End Function

Private Function LOG_FormatLine(ByVal nLevel As Long, _
                                 ByVal sContext As String, _
                                 ByVal sMessage As String) As String
    Dim sTs As String
    sTs = Format$(Now, "yyyy-mm-dd HH:mm:ss")

    If LOG_IsJsonMode() Then
        Dim sObj As String
        sObj = "{"
        sObj = sObj & """ts"":""" & sTs & ""","
        sObj = sObj & """level"":""" & LOG_LevelName(nLevel) & ""","
        sObj = sObj & """ctx"":""" & LOG_JsonEscape(sContext) & ""","
        sObj = sObj & """msg"":""" & LOG_JsonEscape(sMessage) & """"
        sObj = sObj & "}"
        LOG_FormatLine = sObj
    Else
        LOG_FormatLine = sTs & "  " & LOG_LevelNamePadded(nLevel) & "  [" & sContext & "]  " & sMessage
    End If
End Function

Private Function LOG_SanitizePII(ByVal s As String) As String
    Dim sOut As String
    Dim sProfile As String
    Dim nPos As Long
    Dim nEnd As Long

    sOut = s

    sProfile = Environ("USERPROFILE")
    If Len(sProfile) > 0 Then
        sOut = Replace(sOut, sProfile, "<userprofile>", 1, -1, vbTextCompare)
    End If

    nPos = 1
    Do
        nPos = InStr(nPos, sOut, "\Users\", vbTextCompare)
        If nPos = 0 Then Exit Do
        nEnd = InStr(nPos + 7, sOut, "\")
        If nEnd = 0 Then Exit Do
        sOut = Left$(sOut, nPos - 1) & "\Users\<user>" & Mid$(sOut, nEnd)
        nPos = nPos + 12
    Loop

    LOG_SanitizePII = sOut
End Function

Private Function LOG_Asciify(ByVal s As String) As String
    Dim i     As Long
    Dim n     As Long
    Dim ch    As String
    Dim nCode As Long
    Dim sOut  As String

    n = Len(s)
    If n = 0 Then
        LOG_Asciify = ""
        Exit Function
    End If

    sOut = ""
    For i = 1 To n
        ch = Mid$(s, i, 1)
        nCode = AscW(ch)
        If nCode < 0 Then nCode = nCode + 65536
        Select Case nCode
            Case 9, 10, 13
                sOut = sOut & ch
            Case 32 To 126
                sOut = sOut & ch
            Case 183
                sOut = sOut & "|"
            Case Else
                sOut = sOut & "?"
        End Select
    Next i
    LOG_Asciify = sOut
End Function

Public Sub LOG_Open()
    If m_bOpen Then Exit Sub
    If m_bInOpen Then Exit Sub

    m_bInOpen = True
    LOG_OpenInternal
    m_bInOpen = False

    If m_bOpen Then
        LOG_Write LOG_INFO, "LOG", "Session start " & Chr(183) & " " & Chr(34) & Now & Chr(34)
        LOG_Write LOG_INFO, "LOG", "Version: " & PRJ_VERSION
    End If
End Sub

Public Sub LOG_Close()
    If Not m_bOpen Then Exit Sub

    LOG_FlushBuffer
    If m_bWriteFailed Then
        LOG_CloseInternal
        m_sLogPath = ""
        m_bInRotate = False
        Exit Sub
    End If

    Dim sLine As String
    sLine = Format$(Now, "yyyy-mm-dd HH:mm:ss") & "  INFO      [LOG]  "
    sLine = sLine & "Session end " & Chr(183) & " " & Chr(34) & Now & Chr(34)
    sLine = LOG_SanitizePII(sLine)
    sLine = LOG_Asciify(sLine)
    On Error Resume Next
    Print #m_nFileNum, sLine
    If Err.Number <> 0 Then m_bWriteFailed = True
    On Error GoTo 0

    LOG_CloseInternal
    m_sLogPath = ""
    m_bInRotate = False
End Sub

Public Sub LOG_CloseAudit()
    If Not m_bAuditOpen Then Exit Sub
    On Error Resume Next
    Close #m_nAuditFileNum
    On Error GoTo 0
    m_bAuditOpen = False
    m_nAuditFileNum = 0
    m_sAuditPath = ""
End Sub

Private Sub LOG_OpenInternal()
    PRJ_EnsureSandbox

    m_sLogPath = PRJ_LogPath()
    m_nFileNum = FreeFile

    On Error Resume Next
    Err.Clear
    Open m_sLogPath For Append As #m_nFileNum

    If Err.Number = 55 Then
        Err.Clear
        Close
        m_bAuditOpen = False
        m_nAuditFileNum = 0
        m_sAuditPath = ""
        m_nFileNum = FreeFile
        Open m_sLogPath For Append As #m_nFileNum
    End If

    If Err.Number <> 0 Then
        Debug.Print "[modLog.LOG_OpenInternal] Open failed for " & _
                    m_sLogPath & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    m_bOpen = True
    m_bWriteFailed = False
    m_nBufCount = 0
    ReDim m_aLogBuf(0 To LOG_FLUSH_EVERY - 1)
End Sub

Private Sub LOG_OpenAuditInternal()
    If m_bAuditOpen Then Exit Sub

    PRJ_EnsureSandbox
    m_sAuditPath = PRJ_LogAuditPath()
    m_nAuditFileNum = FreeFile

    On Error Resume Next
    Err.Clear
    Open m_sAuditPath For Append As #m_nAuditFileNum
    If Err.Number <> 0 Then
        Debug.Print "[modLog.LOG_OpenAuditInternal] Open failed for " & _
                    m_sAuditPath & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    m_bAuditOpen = True
End Sub

Private Sub LOG_CloseInternal()
    On Error Resume Next
    If m_bOpen Then Close #m_nFileNum
    On Error GoTo 0

    m_bOpen = False
    m_nFileNum = 0
    m_nBufCount = 0
End Sub

Private Sub LOG_FlushBuffer()
    If Not m_bOpen Then
        m_nBufCount = 0
        Exit Sub
    End If
    If m_nBufCount = 0 Then Exit Sub
    If m_bWriteFailed Then
        m_nBufCount = 0
        Exit Sub
    End If

    Dim i As Long
    On Error Resume Next
    For i = 0 To m_nBufCount - 1
        Print #m_nFileNum, m_aLogBuf(i)
        If Err.Number <> 0 Then
            m_bWriteFailed = True
            Err.Clear
            Exit For
        End If
    Next i
    On Error GoTo 0

    m_nBufCount = 0
End Sub

Private Sub LOG_Write(ByVal nLevel As Long, ByVal sContext As String, ByVal sMessage As String)
    If nLevel < g_nLogLevel Then
        ' Even if below the filter, AUDIT contexts still get written
        ' to the audit file. Audit is not user-configurable noise; it
        ' is a compliance record.
        If sContext = "AUDIT" Then LOG_WriteAuditRecord nLevel, sContext, sMessage
        Exit Sub
    End If

    If Not m_bOpen Then
        LOG_Open
        If Not m_bOpen Then Exit Sub
    End If

    If m_bWriteFailed Then Exit Sub

    Dim sLine As String
    sLine = LOG_FormatLine(nLevel, sContext, sMessage)
    sLine = LOG_SanitizePII(sLine)
    sLine = LOG_Asciify(sLine)

    If m_nBufCount >= LOG_FLUSH_EVERY Then
        LOG_FlushBuffer
        If m_bWriteFailed Then Exit Sub
    End If
    m_aLogBuf(m_nBufCount) = sLine
    m_nBufCount = m_nBufCount + 1

    ' Mirror AUDIT records to the audit file regardless of filter.
    If sContext = "AUDIT" Then LOG_WriteAuditRecord nLevel, sContext, sMessage

    If m_nBufCount >= LOG_FLUSH_EVERY And Not m_bInRotate Then
        LOG_FlushBuffer
        If m_bWriteFailed Then Exit Sub
        m_bInRotate = True
        LOG_Rotate
        m_bInRotate = False
    End If
End Sub

' LOG_WriteAuditRecord   append one AUDIT line to the dedicated
'   audit file. Audit records bypass the LogLevel filter (they are
'   compliance records, not diagnostics). Uses the same formatter
'   as the main log so JSON mode applies to both.
'
' v9.19.3: silent Err.Clear replaced with Debug.Print. Cannot use
'   PRJ_LogDebug — that would re-enter LOG_Write from inside the
'   audit write path and potentially trigger LOG_Rotate.
Private Sub LOG_WriteAuditRecord(ByVal nLevel As Long, _
                                  ByVal sContext As String, _
                                  ByVal sMessage As String)
    If Not m_bAuditOpen Then LOG_OpenAuditInternal
    If Not m_bAuditOpen Then Exit Sub

    Dim sLine As String
    sLine = LOG_FormatLine(nLevel, sContext, sMessage)
    sLine = LOG_SanitizePII(sLine)
    sLine = LOG_Asciify(sLine)

    On Error Resume Next
    Err.Clear
    Print #m_nAuditFileNum, sLine
    If Err.Number <> 0 Then
        Debug.Print "[modLog.LOG_WriteAuditRecord] Audit write failed: " & _
                    Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub LOG_WriteError(ByVal sContext As String, ByVal sMessage As String)
    Dim sMsg As String
    sMsg = sMessage
    If Len(g_sCallStack) > 0 Then
        sMsg = sMsg & "  [stack: " & g_sCallStack & "]"
    End If
    LOG_Write LOG_ERROR, sContext, sMsg
End Sub

Private Sub LOG_WriteWarn(ByVal sContext As String, ByVal sMessage As String)
    Dim sMsg As String
    sMsg = sMessage
    If Len(g_sCallStack) > 0 Then
        sMsg = sMsg & "  [stack: " & g_sCallStack & "]"
    End If
    LOG_Write LOG_WARN, sContext, sMsg
End Sub

Private Sub LOG_WriteCritical(ByVal sContext As String, ByVal sMessage As String)
    Dim sMsg As String
    sMsg = sMessage
    If Len(g_sCallStack) > 0 Then
        sMsg = sMsg & "  [stack: " & g_sCallStack & "]"
    End If
    LOG_Write LOG_CRITICAL, sContext, sMsg
End Sub

Public Sub PRJ_LogDebug(ByVal sContext As String, ByVal sMessage As String)
    LOG_Write LOG_DEBUG, sContext, sMessage
End Sub

Public Sub PRJ_LogInfo(ByVal sContext As String, ByVal sMessage As String)
    LOG_Write LOG_INFO, sContext, sMessage
End Sub

Public Sub PRJ_LogWarn(ByVal sContext As String, ByVal sMessage As String)
    LOG_WriteWarn sContext, sMessage
End Sub

Public Sub PRJ_LogError(ByVal sContext As String, ByVal sMessage As String)
    LOG_WriteError sContext, sMessage
End Sub

Public Sub PRJ_LogCritical(ByVal sContext As String, ByVal sMessage As String)
    LOG_WriteCritical sContext, sMessage
End Sub

' PRJ_LogReloadLevel   re-read LogLevel from registry. Call this
'   after changing the key to apply without restart. Called by
'   PRJ_ReloadLogLevel in modEntry.
Public Sub PRJ_LogReloadLevel()
    Dim nNew As Long
    On Error Resume Next
    Err.Clear
    nNew = CLng(REG_Get(REG_KEY_LOGLEVEL, CStr(g_nLogLevel)))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    If nNew < LOG_DEBUG Then nNew = LOG_DEBUG
    If nNew > LOG_CRITICAL Then nNew = LOG_CRITICAL
    g_nLogLevel = nNew
    PRJ_LogInfo "LOG", "Log level reloaded: " & nNew
End Sub

' PRJ_LogReloadJson   re-read LogJson mode from registry.
Public Sub PRJ_LogReloadJson()
    m_bJsonProbed = False
    Dim bNew As Boolean
    bNew = LOG_IsJsonMode()
    PRJ_LogInfo "LOG", "JSON mode reloaded: " & CStr(bNew)
End Sub

Public Sub PRJ_LogFlush()
    If m_bOpen Then LOG_Close
    LOG_CloseAudit
End Sub

Private Sub LOG_Rotate()
    If Not m_bOpen Then Exit Sub

    LOG_FlushBuffer
    If m_bWriteFailed Then Exit Sub

    Dim oFSO As Object
    Set oFSO = PRJ_FSO()

    Dim nSize As Double
    On Error Resume Next
    nSize = oFSO.GetFile(m_sLogPath).Size
    On Error GoTo 0

    If nSize < LOG_MAX_BYTES Then
        Set oFSO = Nothing
        Exit Sub
    End If

    LOG_CloseInternal

    ' Shift existing archives: prev_{N} -> prev_{N+1}, oldest dropped.
    Dim i As Long
    Dim sSrc As String, sDst As String
    For i = LOG_ROTATE_MAX - 1 To 1 Step -1
        sSrc = LOG_GetArchivePath(i)
        sDst = LOG_GetArchivePath(i + 1)
        On Error Resume Next
        Err.Clear
        If oFSO.FileExists(sDst) Then oFSO.DeleteFile sDst, True
        If Err.Number = 0 Then
            If oFSO.FileExists(sSrc) Then oFSO.MoveFile sSrc, sDst
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    Next i

    ' Rotate live log into prev_1.
    Dim sArchive1 As String
    sArchive1 = LOG_GetArchivePath(1)

    On Error Resume Next
    Err.Clear
    If oFSO.FileExists(sArchive1) Then oFSO.DeleteFile sArchive1, True
    If Err.Number = 0 Then
        oFSO.MoveFile m_sLogPath, sArchive1
    End If
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set oFSO = Nothing
        LOG_OpenInternal
        Exit Sub
    End If
    On Error GoTo 0

    Set oFSO = Nothing
    LOG_OpenInternal
End Sub

' END: modLog v9.19.9

