Attribute VB_Name = "modSandbox"
Option Explicit
' Module    : modSandbox
' Purpose   : Sandbox folder lifecycle (path, create, orphan cleanup),
'             path combination helper, canonical text file writer.
' Reference : Appendix A.13, A.24
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   No functional changes.
'
' v9.19.3 (§10.5 Duplicate logic):
'   Added PRJ_WriteTextFile — the canonical text file writer used
'   by modError.Write_CrashReport, modState.STATE_BackupPreferences,
'   and modReportExport.PRJ_ExportReport. Previously each caller
'   had its own FSO.CreateTextFile logic, and modReportExport had
'   a private ADO+FSO fallback pattern. Now: one helper, ADO UTF-8
'   preferred, FSO UTF-16LE fallback. modHelp keeps its own
'   FSO Unicode=True writer because that path is a documented
'   CJK-corruption fix from v9.13.5 and should not be touched.
'
' v9.18.0 (§8.11 File system paths):
'   Added PRJ_PathCombine — canonical path join using FSO.BuildPath
'   with fallback for when FSO is unavailable.
' v9.18.0: Trim -> Trim$, Left/Right -> Left$/Right$ in PRJ_TempRoot.
' v9.13.3 (Sec.1.14): PRJ_EnsureSandbox EH consults PRJ_LogSubsystemBusy().
'
' Reentrancy Guard (this module):
'   m_bInEnsureSandbox — breaks cycle 2. See ARCHITECTURE.md section 14.
'
' Requires  : TEMP_FILE_PREFIX, LOG_FILE_PREFIX, ORPHAN_MAX_AGE_HOURS,
'             ORPHAN_BAK_MAX_AGE_DAYS (modState).
'             PRJ_FSO (modCOM).
'             PRJ_LogError, PRJ_LogSubsystemBusy, PRJ_LogDebug (modLog).
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private m_bInEnsureSandbox As Boolean

Public Function PRJ_SandboxPath() As String
    PRJ_SandboxPath = PRJ_TempRoot() & "\VBADR"
End Function

Public Function PRJ_TempRoot() As String
    Dim s As String
    s = Environ("TEMP")
    If Len(s) = 0 Then s = Environ("TMP")
    If Len(s) = 0 Then s = "C:\Temp"
    Do While Len(s) > 0 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    PRJ_TempRoot = s
End Function

Public Function PRJ_PathCombine(ByVal sBase As String, ByVal sChild As String) As String
    ' §8.11: Canonical path join. Uses FSO.BuildPath when available,
    ' falls back to manual concatenation when FSO creation failed.
    If Len(sBase) = 0 Then
        PRJ_PathCombine = sChild
        Exit Function
    End If
    If Len(sChild) = 0 Then
        PRJ_PathCombine = sBase
        Exit Function
    End If

    Dim oFSO As Object
    On Error Resume Next
    Err.Clear
    Set oFSO = PRJ_FSO()
    If Err.Number <> 0 Then
        Err.Clear
        Set oFSO = Nothing
    End If
    On Error GoTo 0

    If oFSO Is Nothing Then
        If Right$(sBase, 1) = "\" Then
            PRJ_PathCombine = sBase & sChild
        Else
            PRJ_PathCombine = sBase & "\" & sChild
        End If
    Else
        On Error Resume Next
        Err.Clear
        PRJ_PathCombine = oFSO.BuildPath(sBase, sChild)
        If Err.Number <> 0 Then
            Err.Clear
            PRJ_PathCombine = sBase & "\" & sChild
        End If
        On Error GoTo 0
    End If
End Function

' PRJ_WriteTextFile   canonical text file writer.
'   Strategy:
'     1. ADODB.Stream with Charset="utf-8" (preferred — smallest,
'        handles CJK and RTL content, no BOM).
'     2. FSO.CreateTextFile(sPath, True, True) — UTF-16LE with BOM
'        (fallback when ADODB is unavailable).
'   Callers are responsible for ensuring the target directory
'   exists (typically PRJ_EnsureSandbox).
' Params  : sPath — String, full file path
'           sText — String, content to write (overwrites existing)
' Returns : Boolean — True on success, False on any failure
' Throws  : (none) — logs DEBUG on failure, no re-raise
Public Function PRJ_WriteTextFile(ByVal sPath As String, _
                                   ByVal sText As String) As Boolean
    PRJ_WriteTextFile = False
    If Len(sPath) = 0 Then Exit Function

    If PRJ_WriteTextFileViaAdo(sPath, sText) Then
        PRJ_WriteTextFile = True
        Exit Function
    End If
    PRJ_WriteTextFile = PRJ_WriteTextFileViaFso(sPath, sText)
End Function

Private Function PRJ_WriteTextFileViaAdo(ByVal sPath As String, _
                                          ByVal sText As String) As Boolean
    PRJ_WriteTextFileViaAdo = False

    Dim oStream As Object
    Dim bOK As Boolean
    bOK = False

    On Error GoTo EH
    Set oStream = CreateObject("ADODB.Stream")
    If oStream Is Nothing Then Exit Function

    oStream.Type = 2
    oStream.Charset = "utf-8"
    oStream.Open
    oStream.WriteText sText
    oStream.SaveToFile sPath, 2
    bOK = True

Cleanup:
    On Error Resume Next
    If Not oStream Is Nothing Then oStream.Close
    Set oStream = Nothing
    On Error GoTo 0
    PRJ_WriteTextFileViaAdo = bOK
    Exit Function

EH:
    PRJ_LogDebug "PRJ_WriteTextFileViaAdo", "ADODB.Stream failed: " & Err.Description
    bOK = False
    Resume Cleanup
End Function

Private Function PRJ_WriteTextFileViaFso(ByVal sPath As String, _
                                          ByVal sText As String) As Boolean
    PRJ_WriteTextFileViaFso = False

    Dim f As Object

    On Error GoTo EH
    Set f = PRJ_FSO().CreateTextFile(sPath, True, True)
    If f Is Nothing Then Exit Function
    f.Write sText
    f.Close
    Set f = Nothing
    PRJ_WriteTextFileViaFso = True
    Exit Function

EH:
    PRJ_LogDebug "PRJ_WriteTextFileViaFso", "FSO write failed: " & Err.Description
    On Error Resume Next
    If Not f Is Nothing Then f.Close
    Set f = Nothing
    On Error GoTo 0
End Function

Public Sub PRJ_EnsureSandbox()
    If m_bInEnsureSandbox Then Exit Sub
    m_bInEnsureSandbox = True

    On Error GoTo EH

    Dim sPath As String
    sPath = PRJ_SandboxPath()

    Dim oFSO As Object
    Set oFSO = PRJ_FSO()
    If oFSO Is Nothing Then
        m_bInEnsureSandbox = False
        Exit Sub
    End If

    If Not oFSO.FolderExists(sPath) Then
        oFSO.CreateFolder sPath
    End If

    m_bInEnsureSandbox = False
    Exit Sub

EH:
    m_bInEnsureSandbox = False
    If Not PRJ_LogSubsystemBusy() Then
        PRJ_LogError "PRJ_EnsureSandbox", "Cannot create " & sPath & ": " & Err.Description
    End If
End Sub

Public Function PRJ_CleanupOrphans() As Long
    Dim sPath       As String
    Dim oFld        As Object
    Dim oF          As Object
    Dim nDeleted    As Long
    Dim sName       As String
    Dim bIsOrphan   As Boolean
    Dim bIsBak      As Boolean
    Dim colToDelete As Object
    Dim v           As Variant

    PRJ_CleanupOrphans = 0
    PRJ_EnsureSandbox

    sPath = PRJ_SandboxPath()

    On Error Resume Next
    Err.Clear
    Set oFld = PRJ_FSO().GetFolder(sPath)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If oFld Is Nothing Then Exit Function

    Set colToDelete = CreateObject("System.Collections.ArrayList")

    On Error Resume Next
    For Each oF In oFld.Files
        sName = LCase$(oF.Name)
        bIsOrphan = False
        bIsBak = False

        If Left$(sName, Len(TEMP_FILE_PREFIX)) = TEMP_FILE_PREFIX Then
            bIsOrphan = True
        ElseIf Left$(sName, Len(LOG_FILE_PREFIX)) = LCase$(LOG_FILE_PREFIX) Then
            bIsOrphan = True
        ElseIf Right$(sName, 4) = ".bak" Then
            bIsBak = True
        End If

        If bIsOrphan Then
            If DateDiff("h", oF.DateLastModified, Now) > ORPHAN_MAX_AGE_HOURS Then
                colToDelete.Add oF.Path
            End If
        ElseIf bIsBak Then
            If DateDiff("d", oF.DateLastModified, Now) > ORPHAN_BAK_MAX_AGE_DAYS Then
                colToDelete.Add oF.Path
            End If
        End If
    Next
    On Error GoTo 0

    nDeleted = 0
    For Each v In colToDelete
        On Error Resume Next
        Err.Clear
        PRJ_FSO().DeleteFile CStr(v), True
        If Err.Number = 0 Then
            nDeleted = nDeleted + 1
        Else
            Err.Clear
        End If
        On Error GoTo 0
    Next

    PRJ_CleanupOrphans = nDeleted
End Function

' END: modSandbox v9.19.9

