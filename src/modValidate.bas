Attribute VB_Name = "modValidate"
Option Explicit
' Module    : modValidate
' Purpose   : Folder validation orchestration.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.7 (Version alignment): Header and footer synced with
'   PRJ_VERSION. No functional changes from v9.19.6.
'
' v9.19.6 (Audit count cache):
'   Added a memoized audit result. SC_ValidateAudit stores the
'   per-extension counts and the two totals (with/without drawing).
'   SC_UpdateSum2 checks the cache first and skips the four-pass
'   PRJ_CollectByExt scan if the folder and subfolder flag match.
'
' v9.19.6 (SC_ValidateDrawings status classifier):
'   Uses STS_Classify against the ST_MISSING / ST_STALE /
'   ST_MISMATCH / ST_OK statuses added to modStatus in v9.19.6.
'
' v9.18.6: SC_ValidateAudit calls SC_UpdateSum2 at the very end.
'          SC_Validate gains bSkipAudit parameter for txtPath_Change.
' v9.18.5: Single writer for lblSum2.
' v9.18.4: Repaint -> DoEvents at scan/audit transition.
' v9.18.2: No button flicker during scan.
' v9.18.1 (§12.2 Reentrancy): SC_Validate entry guard.
'
' Requires  : g_bInProcessLoop, g_bCancel, g_bAuditMode, g_bSubfolders,
'             g_sLastFolder, MAX_PATH_LEN, SCAN_PROGRESS_MIN_INTERVAL_SEC,
'             MIN_FREE_DISK_BYTES (modState / modConstants).
'             TBadgeCounts (modUIHelpers).
'             SC_SetLabelText, SC_ShowSum2, SC_HideSum2,
'             SC_ApplyDrawingBadge, SC_UpdateSum2 (modUIHelpers).
'             FW_ApplyRunState (modTheme).
'             FIT_FitLabel, FIT_ELL_MIDDLE, FIT_ELL_END (modFIT).
'             PRJ_AlignTxt (modTheme).
'             SN_SanitizePath (modPath).
'             PRJ_FSO (modCOM).
'             PRJ_CheckDiskSpace (modFileInfo).
'             PRJ_CollectIDW (modBatch).
'             PRJ_CheckPDFStatus, PRJ_HasDrawing (modInspect).
'             PRJ_GetAuditSummary (modSummary).
'             PRJ_CollectByExt (modSummary).
'             TM_Now, TM_Elapsed (modTime).
'             L, L2, L3 (modLang).
'             PRJ_LogError, PRJ_LogWarn (modLog).
'             STS_Classify, TStatus, ST_* (modStatus).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private m_bAuditValid    As Boolean
Private m_sAuditFolder   As String
Private m_bAuditSub      As Boolean
Private m_nAuditIam      As Long
Private m_nAuditIpt      As Long
Private m_nAuditWith     As Long
Private m_nAuditWithout  As Long

Public Function PRJ_ValidateFolderPath(ByVal oFrm As Object, _
                                        ByVal bRequireDiskSpace As Boolean) As Boolean
    PRJ_ValidateFolderPath = False

    If Not SC_QuickCheck(oFrm) Then Exit Function

    If bRequireDiskSpace Then
        Dim sPath As String
        sPath = SN_SanitizePath(oFrm.Controls("txtPath").Text)

        Dim dFree As Double
        If Not PRJ_CheckDiskSpace(sPath, MIN_FREE_DISK_BYTES, dFree) Then
            Dim oSum As Object
            Set oSum = oFrm.Controls("lblSum")
            If Not oSum Is Nothing Then
                SC_SetLabelText oSum, _
                    L2("scan.disk.insufficient", PRJ_FormatFileSize(dFree)), _
                    CLR_ERROR
            End If
            SC_HideSum2 oFrm
            FW_ApplyRunState oFrm, False
            PRJ_LogWarn "PRJ_ValidateFolderPath", _
                "Insufficient disk space at " & sPath & ": " & dFree & " bytes free"
            Exit Function
        End If
    End If

    PRJ_ValidateFolderPath = True
End Function

Public Function SC_QuickCheck(oFrm As Object) As Boolean
    SC_QuickCheck = False
    On Error GoTo EH

    Dim oSum  As Object
    Dim sPath As String

    Set oSum = oFrm.Controls("lblSum")
    sPath = SN_SanitizePath(oFrm.Controls("txtPath").Text)

    If Len(sPath) = 0 Then
        SC_SetLabelText oSum, L("scan.no.folder"), CLR_ERROR
        SC_HideSum2 oFrm
        FW_ApplyRunState oFrm, False
        Exit Function
    End If

    If Len(sPath) > MAX_PATH_LEN Then
        SC_SetLabelText oSum, L2("scan.path.too.long", CStr(MAX_PATH_LEN)), CLR_ERROR
        SC_HideSum2 oFrm
        FW_ApplyRunState oFrm, False
        Exit Function
    End If

    If Not PRJ_FSO().FolderExists(sPath) Then
        SC_SetLabelText oSum, L("scan.no.folder.exists"), CLR_ERROR
        SC_HideSum2 oFrm
        FW_ApplyRunState oFrm, False
        Exit Function
    End If

    SC_SetLabelText oSum, L("scan.path.ok"), CLR_WARN
    SC_HideSum2 oFrm
    FW_ApplyRunState oFrm, True
    SC_QuickCheck = True
    Exit Function

EH:
    PRJ_LogError "SC_QuickCheck", Err.Description
    SC_QuickCheck = False
End Function

Public Sub SC_Validate(oFrm As Object, Optional ByVal bSkipAudit As Boolean = False)
    If g_bInProcessLoop Then Exit Sub

    On Error GoTo EH

    If Not SC_QuickCheck(oFrm) Then Exit Sub

    Dim sPath As String
    sPath = SN_SanitizePath(oFrm.Controls("txtPath").Text)

    Dim oSum As Object
    Set oSum = oFrm.Controls("lblSum")

    Dim bWillRefreshAudit As Boolean
    bWillRefreshAudit = (Not bSkipAudit) And g_bAuditMode
    If bWillRefreshAudit Then m_bAuditValid = False

    SC_SetLabelText oSum, L("scan.collecting"), CLR_TEXT_SEC
    SC_HideSum2 oFrm
    DoEvents

    g_bCancel = False
    g_bInProcessLoop = True

    Dim colIdw As Object
    Set colIdw = CreateObject("System.Collections.ArrayList")
    PRJ_CollectIDW sPath, colIdw

    If g_bCancel Then GoTo Cancelled

    Dim nTot As Long
    nTot = colIdw.Count

    If nTot = 0 Then
      g_bInProcessLoop = False
      SC_SetLabelText oSum, L("scan.sum.no.idw"), CLR_ERROR
      SC_HideSum2 oFrm
       FW_ApplyRunState oFrm, False
        Exit Sub
    End If

    Dim tB As TBadgeCounts
    tB.nTot = nTot

    SC_ValidateDrawings colIdw, tB, oFrm

    If g_bCancel Then GoTo Cancelled

    SC_ApplyDrawingBadge oFrm, tB

    DoEvents

    If bSkipAudit Then
        SC_HideSum2 oFrm
    Else
        If g_bAuditMode Then
            SC_ValidateAudit oFrm, sPath
        Else
            SC_HideSum2 oFrm
        End If
    End If

    If g_bCancel Then GoTo Cancelled

    g_bInProcessLoop = False
    g_bCancel = False
    FW_ApplyRunState oFrm, True
    Exit Sub

Cancelled:
    g_bInProcessLoop = False
    g_bCancel = False
    SC_SetLabelText oSum, L("scan.cancelled"), CLR_WARN
    SC_HideSum2 oFrm
    FW_ApplyRunState oFrm, False
    Exit Sub

EH:
    g_bInProcessLoop = False
    g_bCancel = False
    PRJ_LogError "SC_Validate", Err.Description
    SC_SetLabelText oFrm.Controls("lblSum"), L("scan.validate.error"), CLR_ERROR
    SC_HideSum2 oFrm
    FW_ApplyRunState oFrm, False
End Sub

Private Function SC_TickProgress(ByVal oFrm As Object, _
                                   ByVal oLbl As Object, _
                                   ByVal sFormatKey As String, _
                                   ByVal nCurrent As Long, _
                                   ByVal nTotal As Long, _
                                   ByRef dLastProg As Double) As Boolean
    SC_TickProgress = False

    Dim dNow As Double
    dNow = TM_Now()
    If TM_Elapsed(dLastProg, dNow) < SCAN_PROGRESS_MIN_INTERVAL_SEC Then Exit Function

    SC_SetLabelText oLbl, L3(sFormatKey, CStr(nCurrent), CStr(nTotal)), CLR_TEXT_SEC
    DoEvents
    dLastProg = dNow
    SC_TickProgress = True
End Function

Private Sub SC_ValidateDrawings(ByRef colIdw As Object, _
                                 ByRef t As TBadgeCounts, _
                                 ByVal oFrm As Object)
    Dim i            As Long
    Dim sStat        As String
    Dim oSum         As Object
    Dim dLastProg    As Double
    Dim nTot         As Long
    Dim st           As TStatus

    On Error GoTo EH

    Set oSum = oFrm.Controls("lblSum")
    nTot = colIdw.Count
    dLastProg = TM_Now()

    For i = 0 To nTot - 1
        If g_bCancel Then Exit For

        On Error Resume Next
        Err.Clear
        sStat = PRJ_CheckPDFStatus(CStr(colIdw(i)))
        If Err.Number <> 0 Then
            Err.Clear
            sStat = "Error"
        End If
        On Error GoTo EH

        st = STS_Classify(sStat)
        Select Case st
            Case ST_ERROR:      t.nErr = t.nErr + 1
            Case ST_MISSING:    t.nMissing = t.nMissing + 1
            Case ST_STALE:      t.nStale = t.nStale + 1
            Case ST_MISMATCH:   t.nMismatch = t.nMismatch + 1
            Case ST_OK:         t.nReady = t.nReady + 1
            Case Else:          t.nErr = t.nErr + 1
        End Select

        SC_TickProgress oFrm, oSum, "scan.analysing", i + 1, nTot, dLastProg
    Next i
    Exit Sub

EH:
    PRJ_LogError "SC_ValidateDrawings", Err.Description
End Sub

Private Sub SC_ValidateAudit(ByVal oFrm As Object, ByVal sPath As String)
    Dim colIam       As Object
    Dim colIpt       As Object
    Dim nModels      As Long
    Dim nHas         As Long
    Dim nNo          As Long
    Dim i            As Long
    Dim oSum2        As Object
    Dim sTip2        As String
    Dim dLastProg    As Double

    On Error GoTo EH

    Set colIam = CreateObject("System.Collections.ArrayList")
    Set colIpt = CreateObject("System.Collections.ArrayList")

    PRJ_CollectByExt sPath, g_bSubfolders, "iam", colIam
    PRJ_CollectByExt sPath, g_bSubfolders, "ipt", colIpt
    PRJ_CollectByExt sPath, g_bSubfolders, "ipn", colIpt

    nModels = colIam.Count + colIpt.Count
    If nModels = 0 Then
        m_bAuditValid = False
        SC_HideSum2 oFrm
        Exit Sub
    End If

    Set oSum2 = oFrm.Controls("lblSum2")
    SC_ShowSum2 oFrm
    SC_SetLabelText oSum2, L("scan.auditing.start"), CLR_TEXT_SEC
    DoEvents

    dLastProg = TM_Now()

    Dim nDone As Long
    nDone = 0

    For i = 0 To colIam.Count - 1
        If g_bCancel Then Exit For
        If PRJ_HasDrawing(CStr(colIam(i))) Then
            nHas = nHas + 1
        Else
            nNo = nNo + 1
        End If
        nDone = nDone + 1
        SC_TickProgress oFrm, oSum2, "scan.auditing", nDone, nModels, dLastProg
    Next i
    For i = 0 To colIpt.Count - 1
        If g_bCancel Then Exit For
        If PRJ_HasDrawing(CStr(colIpt(i))) Then
            nHas = nHas + 1
        Else
            nNo = nNo + 1
        End If
        nDone = nDone + 1
        SC_TickProgress oFrm, oSum2, "scan.auditing", nDone, nModels, dLastProg
    Next i

    If g_bCancel Then Exit Sub

    m_sAuditFolder = sPath
    m_bAuditSub = g_bSubfolders
    m_nAuditIam = colIam.Count
    m_nAuditIpt = colIpt.Count
    m_nAuditWith = nHas
    m_nAuditWithout = nNo
    m_bAuditValid = True

    sTip2 = L3("scan.models.detail", CStr(nHas), CStr(nNo))
    oSum2.ControlTipText = sTip2

    SC_UpdateSum2 oFrm
    Exit Sub

EH:
    PRJ_LogError "SC_ValidateAudit", Err.Description
    m_bAuditValid = False
End Sub

Public Function SC_GetCachedAudit(ByVal sFolder As String, _
                                   ByVal bSubfolders As Boolean, _
                                   ByRef nIam As Long, _
                                   ByRef nIpt As Long, _
                                   ByRef nWith As Long, _
                                   ByRef nWithout As Long) As Boolean
    SC_GetCachedAudit = False
    If Not m_bAuditValid Then Exit Function
    If LCase$(sFolder) <> LCase$(m_sAuditFolder) Then Exit Function
    If bSubfolders <> m_bAuditSub Then Exit Function

    nIam = m_nAuditIam
    nIpt = m_nAuditIpt
    nWith = m_nAuditWith
    nWithout = m_nAuditWithout
    SC_GetCachedAudit = True
End Function

' END: modValidate v9.19.9

