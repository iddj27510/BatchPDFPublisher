Attribute VB_Name = "modBatch"
Option Explicit
' Module    : modBatch
' Purpose   : Batch orchestration over a folder of IDW files.
' Reference : Appendix A.13, A.31
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + Requires fix):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Requires section corrected: PRJ_GetRelPath (modFileInfo) is
'   NOT a dependency of this module — it is mentioned only in the
'   v9.19.3 changelog comment (the report list uses it, but that
'   happens in modReportBuild, not here). Removed from the list.
'   No functional changes.
'
' v9.19.3 (Full path in progress label):
'   PRJ_ScanFolder now stores the full IDW path in g_sCurrentRelPath
'   and passes it to PRJ_UpdateProgress. frmProgress.lblFile shows
'   the full path (via caption + ControlTipText). The report list
'   still uses PRJ_GetRelPath for its relative-path column, so the
'   two views are independent.
'   Variable name g_sCurrentRelPath kept for compatibility but now
'   carries the absolute path. See modState header.
'
' v9.18.0 (§1.3 refactor):
'   Added PRJ_RunScanSession. frmProgress.UserForm_Activate now calls
'   this single procedure instead of calling PRJ_ScanFolder and setting
'   g_dTimeEnd inline.
'
' Requires  : g_colFiles, g_colLog, g_nTotal, g_nDone, g_bCancel,
'             g_sCurrentRelPath, g_sLastFolder, g_bSubfolders,
'             g_bForceOverwrite, g_dTimeEnd (modState).
'             PRJ_FSO (modCOM).
'             PRJ_CreatePDF (modPDFExport).
'             PRJ_UpdateProgress (modPDFProgress).
'             PRJ_CollectByExt (modSummary).
'             TM_Now (modTime).
'             L (modLang).
'             PRJ_LogError (modLog).
'
' Notes     : No line-continuation characters in this module.
'             No decorative separator comment lines.

Public Sub PRJ_InitCollections()
    Set g_colFiles = CreateObject("System.Collections.ArrayList")

    Set g_colLog = CreateObject("Scripting.Dictionary")
    g_colLog.CompareMode = vbTextCompare
End Sub

Private Sub PRJ_EnsureCollections()
    If g_colFiles Is Nothing Then
        Set g_colFiles = CreateObject("System.Collections.ArrayList")
    End If
    If g_colLog Is Nothing Then
        Set g_colLog = CreateObject("Scripting.Dictionary")
        g_colLog.CompareMode = vbTextCompare
    End If
End Sub

Public Sub PRJ_CollectIDW(ByVal sFolder As String, ByVal colOut As Object)
    If colOut Is Nothing Then Exit Sub
    On Error Resume Next
    PRJ_CollectByExt sFolder, g_bSubfolders, "idw", colOut
    On Error GoTo 0
End Sub

Public Sub PRJ_CollectModels(ByVal sFolder As String, ByVal colOut As Object)
    If colOut Is Nothing Then Exit Sub
    On Error Resume Next
    PRJ_CollectByExt sFolder, g_bSubfolders, "ipt", colOut
    PRJ_CollectByExt sFolder, g_bSubfolders, "iam", colOut
    PRJ_CollectByExt sFolder, g_bSubfolders, "ipn", colOut
    On Error GoTo 0
End Sub

Public Sub PRJ_RunScanSession()
    On Error GoTo EH

    PRJ_ScanFolder g_sLastFolder
    g_dTimeEnd = TM_Now()
    Exit Sub

EH:
    PRJ_LogError "PRJ_RunScanSession", Err.Description
    g_dTimeEnd = TM_Now()
End Sub

Public Sub PRJ_ScanFolder(ByVal sFolder As String)
    If Len(sFolder) = 0 Then Exit Sub

    Dim bFolderExists As Boolean
    On Error Resume Next
    Err.Clear
    bFolderExists = PRJ_FSO().FolderExists(sFolder)
    If Err.Number <> 0 Then
        Err.Clear
        bFolderExists = False
    End If
    On Error GoTo EH
    If Not bFolderExists Then
        PRJ_LogError "PRJ_ScanFolder", "Folder does not exist: " & sFolder
        Exit Sub
    End If

    PRJ_EnsureCollections

    Dim colIdw As Object
    Set colIdw = CreateObject("System.Collections.ArrayList")
    PRJ_CollectIDW sFolder, colIdw

    Dim nTot As Long
    nTot = colIdw.Count
    g_nTotal = nTot
    g_nDone = 0

    If nTot = 0 Then
        g_sCurrentRelPath = ""
        Exit Sub
    End If

    Dim i       As Long
    Dim sIDW    As String
    Dim sStat   As String

    For i = 0 To nTot - 1
        If g_bCancel Then Exit For

        sIDW = CStr(colIdw(i))
        ' v9.19.3: g_sCurrentRelPath now carries the full path so
        ' frmProgress.lblFile can display it. The report list still
        ' computes its own relative path via PRJ_GetRelPath.
        g_sCurrentRelPath = sIDW

        On Error Resume Next
        PRJ_UpdateProgress sIDW, L("progress.processing")
        On Error GoTo EH

        On Error Resume Next
        Err.Clear
        sStat = PRJ_CreatePDF(sIDW, g_bForceOverwrite)
        If Err.Number <> 0 Then
            Dim sErrDesc As String
            sErrDesc = Err.Description
            Err.Clear
            On Error GoTo EH
            sStat = "Error: " & sErrDesc
            PRJ_LogError "PRJ_ScanFolder", "[" & sIDW & "] " & sErrDesc
        End If
        On Error GoTo EH

        g_colFiles.Add Array(sIDW, sStat)
        If Not g_colLog Is Nothing Then
            If Not g_colLog.Exists(sIDW) Then
                g_colLog.Add sIDW, sStat
            Else
                g_colLog(sIDW) = sStat
            End If
        End If

        g_nDone = i + 1

        On Error Resume Next
        PRJ_UpdateProgress sIDW, sStat
        On Error GoTo EH

        DoEvents
    Next i

    g_sCurrentRelPath = ""
    Exit Sub

EH:
    g_sCurrentRelPath = ""
    PRJ_LogError "PRJ_ScanFolder", Err.Description
End Sub

' END: modBatch v9.19.9
