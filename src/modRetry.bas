Attribute VB_Name = "modRetry"
Option Explicit
' Module    : modRetry
' Purpose   : File operation retry helpers with Full Jitter.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + header standardization):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Header converted from Style A (free-form) to Style B
'   (Module/Purpose/Reference/Version). Added missing Purpose
'   and Reference fields.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in RETRY_DeleteFile, RETRY_MoveFile and RETRY_Pause
'   signatures). No functional changes.
'
' v9.18.0 (§2.11 Magic numbers):
'   RETRY_Pause now uses UI_SLEEP_CHUNK_MS instead of literal 50.
' v9.18.0: Err.Clear before each FSO call (stale-Err hygiene,
'          matching modClipboard / modCOM / modInspect v9.18.0).
'          Header sync.
' v9.14.8 (Faza 4, C.7): RETRY_Pause uses TM_Now / TM_Elapsed.
' v9.13.9: removed local RETRY_* constants - now in modState.
'
' Requires  : RETRY_MAX_ATTEMPTS, RETRY_BASE_SEC, RETRY_BACKOFF_MULT,
'             RETRY_MIN_MS, MS_PER_SECOND, UI_SLEEP_CHUNK_MS (modState).
'             TM_Now / TM_Elapsed (modTime), PRJ_FSO (modCOM),
'             PRJ_LogWarn / PRJ_LogError (modLog).
'             Sleep (modWinAPI).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Function RETRY_DeleteFile(sPath As String, _
                                  Optional nMaxAttempts As Long = RETRY_MAX_ATTEMPTS) As Boolean
    Dim i      As Long
    Dim dDelay As Double
    Dim bOK    As Boolean

    RETRY_DeleteFile = False

    For i = 1 To nMaxAttempts
        On Error Resume Next
        Err.Clear
        PRJ_FSO().DeleteFile sPath, True
        bOK = (Err.Number = 0)
        If Not bOK Then Err.Clear
        On Error GoTo 0

        If bOK Then
            RETRY_DeleteFile = True
            Exit Function
        End If

        If i < nMaxAttempts Then
            dDelay = RETRY_BASE_SEC * (RETRY_BACKOFF_MULT ^ (i - 1))
            PRJ_LogWarn "RETRY_DeleteFile", "Attempt " & i & " failed for " & sPath
            RETRY_Pause dDelay
        End If
    Next i

    PRJ_LogError "RETRY_DeleteFile", "Gave up: " & sPath
End Function

Public Function RETRY_MoveFile(sFrom As String, sTo As String, _
                                Optional nMaxAttempts As Long = RETRY_MAX_ATTEMPTS) As Boolean
    Dim i      As Long
    Dim dDelay As Double
    Dim bOK    As Boolean

    RETRY_MoveFile = False

    For i = 1 To nMaxAttempts
        On Error Resume Next
        Err.Clear
        PRJ_FSO().MoveFile sFrom, sTo
        bOK = (Err.Number = 0)
        If Not bOK Then Err.Clear
        On Error GoTo 0

        If bOK Then
            RETRY_MoveFile = True
            Exit Function
        End If

        If i < nMaxAttempts Then
            dDelay = RETRY_BASE_SEC * (RETRY_BACKOFF_MULT ^ (i - 1))
            PRJ_LogWarn "RETRY_MoveFile", "Attempt " & i & " failed for " & sFrom
            RETRY_Pause dDelay
        End If
    Next i

    PRJ_LogError "RETRY_MoveFile", "Gave up: " & sFrom
End Function

Private Sub RETRY_Pause(dSeconds As Double)
    Dim dJittered As Double
    Dim dStart    As Double
    Dim nFloorMs  As Long

    dJittered = Rnd() * dSeconds

    nFloorMs = RETRY_MIN_MS
    If dJittered * MS_PER_SECOND < nFloorMs Then
        dJittered = nFloorMs / MS_PER_SECOND
    End If

    dStart = TM_Now()

    Do While TM_Elapsed(dStart, TM_Now()) < dJittered
        DoEvents
        Sleep UI_SLEEP_CHUNK_MS
    Loop
End Sub

' END: modRetry v9.19.9

