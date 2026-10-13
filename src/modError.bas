Attribute VB_Name = "modError"
Option Explicit
' Module    : modError
' Purpose   : Centralized error reporting, classification, and
'             application-level handler
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Documentation accuracy):
'   Removed duplicate changelog entry (v9.19.6 was listed twice).
'   Requires section corrected: EC_*, LOG_* and PRJ_VERSION live
'   in modConstants, not modState; only g_sCallStack is in
'   modState. Added PRJ_LogCritical (used by PRJ_Fatal) and
'   PRJ_LogDebug (used by Write_CrashReport), both previously
'   missing.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in procedure signatures and in multi-line
'   PRJ_LogError / PRJ_PathCombine calls).
'   No functional changes.
'
' v9.19.3 (§10.5 Duplicate logic):
'   Write_CrashReport now uses modSandbox.PRJ_WriteTextFile instead
'   of its own FSO.CreateTextFile(sFile, True, True). One writer,
'   one code path.
' v9.18.0 (§8.11 File system paths):
'   Write_CrashReport uses PRJ_PathCombine for crash file path.
' v9.18.0 (§4.x Error handling):
'   PRJ_ShowError labels localized via L("errdlg.label.*").
'   Added ERR_Classify, ERR_IsRecoverable, ERR_DescribeClass,
'   ERR_GlobalHandler.
'   PRJ_Fatal labels remain EN (A.66 debuggability).
'
' Requires  : g_sCallStack (modState).
'             EC_*, LOG_*, PRJ_VERSION (modConstants).
'             L, L2, LN (modLang).
'             PRJ_LogError, PRJ_LogCritical, PRJ_LogDebug,
'             PRJ_LogFlush, PRJ_LogPath, PRJ_LogWriteFailed,
'             LOG_Open (modLog).
'             PRJ_EnsureSandbox, PRJ_SandboxPath, PRJ_PathCombine,
'             PRJ_WriteTextFile (modSandbox).
'             PRJ_FSO (modCOM).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-argument call
'             continuation are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Function PRJ_ShowError(sContext As String, sMessage As String, Optional bAllowRetry As Boolean = False) As Long
    Dim sTitle    As String
    Dim sStack    As String
    Dim sLogRef   As String
    Dim sTime     As String
    Dim sBody     As String
    Dim nResult   As Long
    Dim sPrompt   As String

    On Error GoTo Safe

    sTitle = L("title") & " - " & sContext

    If Len(g_sCallStack) > 0 Then
        sStack = vbCrLf & L("errdlg.label.stack") & ": " & g_sCallStack
    Else
        sStack = ""
    End If

    sLogRef = PRJ_LogPath()
    sTime = Format$(Now, "dd.MM.yyyy HH:mm:ss")

    sBody = sMessage & vbCrLf & vbCrLf
    sBody = sBody & L("errdlg.label.context") & ": " & sContext & sStack & vbCrLf
    sBody = sBody & L("errdlg.label.time") & ": " & sTime & vbCrLf
    sBody = sBody & L("errdlg.label.log") & ": " & sLogRef

    PRJ_LogError sContext, sMessage

    If bAllowRetry Then
        sPrompt = sBody & vbCrLf & vbCrLf & L("errdlg.retry.prompt")
        nResult = MsgBox(sPrompt, vbAbortRetryIgnore + vbExclamation, sTitle)
        Select Case nResult
            Case vbRetry:  PRJ_ShowError = 1
            Case vbIgnore: PRJ_ShowError = 2
            Case vbAbort:  PRJ_ShowError = 3
            Case Else:     PRJ_ShowError = 0
        End Select
    Else
        MsgBox sBody, vbOKOnly + vbExclamation, sTitle
        PRJ_ShowError = 0
    End If
    Exit Function

Safe:
    On Error Resume Next
    MsgBox sContext & ": " & sMessage, vbExclamation, "Batch PDF Publisher"
    PRJ_ShowError = 0
End Function

Public Sub PRJ_Fatal(sContext As String, sMessage As String)
    Dim sFatalLog  As String
    Dim sCrashRpt  As String
    Dim sFatalBody As String
    Dim bHaveCrash As Boolean
    Dim bLogFailed As Boolean
    Dim sTitle     As String

    On Error GoTo Safe

    PRJ_LogCritical sContext, "FATAL: " & sMessage

    On Error Resume Next
    sCrashRpt = Write_CrashReport(sContext, sMessage)
    bHaveCrash = (Len(sCrashRpt) > 0)
    Err.Clear
    On Error GoTo Safe

    PRJ_LogFlush

    bLogFailed = PRJ_LogWriteFailed()
    If bLogFailed Then
        LOG_Open
    End If

    sFatalLog = PRJ_LogPath()
    sTitle = L("title")

    sFatalBody = "Fatal error in " & sContext & ":" & vbCrLf & vbCrLf
    sFatalBody = sFatalBody & sMessage & vbCrLf & vbCrLf
    If bHaveCrash Then
        sFatalBody = sFatalBody & "Crash report: " & sCrashRpt & vbCrLf
    End If
    sFatalBody = sFatalBody & "See log: " & sFatalLog
    If bLogFailed Then
        sFatalBody = sFatalBody & " (log write failed - crash report is authoritative)"
    End If

    MsgBox sFatalBody, vbCritical, sTitle
    Exit Sub

Safe:
    On Error Resume Next
    PRJ_LogFlush
    MsgBox "FATAL: " & sContext & " - " & sMessage, vbCritical, "Batch PDF Publisher"
End Sub

Public Function ERR_Classify(ByVal nErrNumber As Long) As Long
    Select Case nErrNumber
        Case 18
            ERR_Classify = EC_USER_CANCEL
        Case 53, 76, 68
            ERR_Classify = EC_NETWORK
        Case 70, 75
            ERR_Classify = EC_PERMISSION
        Case 9, 13, 94
            ERR_Classify = EC_DATA
        Case 55, 58, 61
            ERR_Classify = EC_FILE_LOCK
        Case 429
            ERR_Classify = EC_COM
        Case -2147024891
            ERR_Classify = EC_PERMISSION
        Case -2147467259
            ERR_Classify = EC_COM
        Case Else
            ERR_Classify = EC_UNKNOWN
    End Select
End Function

Public Function ERR_IsRecoverable(ByVal nClass As Long) As Boolean
    Select Case nClass
        Case EC_NETWORK, EC_FILE_LOCK, EC_USER_CANCEL
            ERR_IsRecoverable = True
        Case Else
            ERR_IsRecoverable = False
    End Select
End Function

Public Function ERR_DescribeClass(ByVal nClass As Long) As String
    Select Case nClass
        Case EC_NETWORK:      ERR_DescribeClass = "network"
        Case EC_PERMISSION:   ERR_DescribeClass = "permission"
        Case EC_DATA:         ERR_DescribeClass = "data"
        Case EC_FILE_LOCK:    ERR_DescribeClass = "file-lock"
        Case EC_COM:          ERR_DescribeClass = "com"
        Case EC_USER_CANCEL:  ERR_DescribeClass = "user-cancel"
        Case Else:            ERR_DescribeClass = "unknown"
    End Select
End Function

Public Sub ERR_GlobalHandler(ByVal sContext As String)
    Dim nClass As Long
    Dim sDesc As String
    Dim sMsg As String

    On Error Resume Next

    sMsg = Err.Description
    If Len(sMsg) = 0 Then sMsg = "Unknown error (err " & Err.Number & ")"

    nClass = ERR_Classify(Err.Number)
    sDesc = ERR_DescribeClass(nClass)

    PRJ_LogError sContext, "Global handler [" & sDesc & "] err=" & Err.Number & ": " & sMsg

    If ERR_IsRecoverable(nClass) Then
        PRJ_ShowError sContext, sMsg, True
    Else
        PRJ_ShowError sContext, sMsg, False
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Private Function Write_CrashReport(ByVal sContext As String, ByVal sMessage As String) As String
    Dim sFile As String
    Dim sBody As String

    Write_CrashReport = ""

    On Error GoTo Fail

    PRJ_EnsureSandbox
    sFile = PRJ_PathCombine(PRJ_SandboxPath(), "crash_" & Format$(Now, "yyyy-mm-dd_hh-nn-ss") & ".txt")

    sBody = "Batch PDF Publisher - Crash Report" & vbCrLf
    sBody = sBody & "Timestamp: " & Format$(Now, "yyyy-mm-dd HH:mm:ss") & vbCrLf
    sBody = sBody & "Version:   " & PRJ_VERSION & vbCrLf
    sBody = sBody & "Context:   " & sContext & vbCrLf
    sBody = sBody & vbCrLf & "Message:" & vbCrLf
    sBody = sBody & sMessage & vbCrLf

    If Len(g_sCallStack) > 0 Then
        sBody = sBody & vbCrLf & "Call stack:" & vbCrLf
        sBody = sBody & g_sCallStack & vbCrLf
    End If

    sBody = sBody & vbCrLf & "Log: " & PRJ_LogPath() & vbCrLf

    ' v9.19.3: unified writer (modSandbox.PRJ_WriteTextFile).
    If Not PRJ_WriteTextFile(sFile, sBody) Then Exit Function

    Write_CrashReport = sFile
    Exit Function

Fail:
    PRJ_LogDebug "Write_CrashReport", "Crash report write failed: " & Err.Description
    Write_CrashReport = ""
End Function

' END: modError v9.19.9

