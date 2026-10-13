Attribute VB_Name = "modSummary"
Option Explicit
' Module    : modSummary
' Purpose   : Audit aggregations for summary 1 and summary 2
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Added missing Reference field for consistency with the rest
'   of the project.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in procedure signatures and in the return-value
'   assignment in PRJ_GetAuditSummary). No functional changes.
'
' v9.19.0 (Audit summary narrative):
'   PRJ_GetAuditSummary now uses LN("summary.hdr.models_full", ...)
'   to produce a single "{0} models ({1} assemblies, {2} parts)"
'   segment instead of three separate ASSY/PART/IDW segments.
'   "with_idw" and "no_idw" reworded to "{0} covered" and
'   "{0} missing" in all 13 languages.
' v9.18.0 (§1.17 Resource Separation):
'   PRJ_GetAuditSummary no longer hardcodes "ASSY:", "PART:", "IDW:",
'   "With .idw:", "No .idw:".
'   Err.Clear before each FSO call (stale-Err hygiene).
'   Per-operation OERN scope in PRJ_CollectByExt.
'
' Requires  : PRJ_FSO (modCOM).
'             L, L2, LN (modLang).
'             PRJ_LogError (modLog).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-argument return-value
'             assignments are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Function PRJ_GetAuditSummary(ByVal sFolder As String, _
                                     ByVal bSubfolders As Boolean) As String
    PRJ_GetAuditSummary = ""

    On Error GoTo EH
    If Len(sFolder) = 0 Then Exit Function

    Dim bFolderExists As Boolean
    On Error Resume Next
    Err.Clear
    bFolderExists = PRJ_FSO().FolderExists(sFolder)
    If Err.Number <> 0 Then
        Err.Clear
        bFolderExists = False
    End If
    On Error GoTo EH
    If Not bFolderExists Then Exit Function

    Dim colIam As Object, colIpt As Object, colIdw As Object
    Set colIam = CreateObject("System.Collections.ArrayList")
    Set colIpt = CreateObject("System.Collections.ArrayList")
    Set colIdw = CreateObject("System.Collections.ArrayList")

    PRJ_CollectByExt sFolder, bSubfolders, "iam", colIam
    PRJ_CollectByExt sFolder, bSubfolders, "ipt", colIpt
    PRJ_CollectByExt sFolder, bSubfolders, "ipn", colIpt
    PRJ_CollectByExt sFolder, bSubfolders, "idw", colIdw

    Dim nIam As Long, nIpt As Long, nIdw As Long
    nIam = colIam.Count
    nIpt = colIpt.Count
    nIdw = colIdw.Count

    Dim oDictIdw As Object
    Set oDictIdw = CreateObject("Scripting.Dictionary")
    oDictIdw.CompareMode = 1

    Dim i As Long
    For i = 0 To colIdw.Count - 1
        oDictIdw(LCase$(PRJ_FSO().GetBaseName(CStr(colIdw(i))))) = True
    Next i

    Dim nHasIdw As Long, nNoIdw As Long
    For i = 0 To colIam.Count - 1
        If oDictIdw.Exists(LCase$(PRJ_FSO().GetBaseName(CStr(colIam(i))))) Then
            nHasIdw = nHasIdw + 1
        Else
            nNoIdw = nNoIdw + 1
        End If
    Next i
    For i = 0 To colIpt.Count - 1
        If oDictIdw.Exists(LCase$(PRJ_FSO().GetBaseName(CStr(colIpt(i))))) Then
            nHasIdw = nHasIdw + 1
        Else
            nNoIdw = nNoIdw + 1
        End If
    Next i

    Dim sep As String
    sep = " " & Chr(183) & " "

    Dim nTotal As Long
    nTotal = nIam + nIpt

    PRJ_GetAuditSummary = _
        LN("summary.hdr.models_full", CStr(nTotal), CStr(nIam), CStr(nIpt)) & sep & _
        L2("summary.hdr.with_idw", CStr(nHasIdw)) & sep & _
        L2("summary.hdr.no_idw", CStr(nNoIdw))
    Exit Function

EH:
    PRJ_GetAuditSummary = ""
End Function

Public Sub PRJ_CollectByExt(ByVal sFolder As String, _
                             ByVal bSub As Boolean, _
                             ByVal sExt As String, _
                             ByVal colOut As Object)
    Dim oFld As Object
    Dim oFile As Object
    Dim oSub As Object

    On Error Resume Next
    Err.Clear
    Set oFld = PRJ_FSO().GetFolder(sFolder)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    If oFld Is Nothing Then Exit Sub

    On Error Resume Next
    For Each oFile In oFld.Files
        If LCase$(PRJ_FSO().GetExtensionName(oFile.Name)) = LCase$(sExt) Then
            colOut.Add oFile.Path
        End If
        DoEvents
    Next oFile
    On Error GoTo 0

    If bSub Then
        On Error Resume Next
        For Each oSub In oFld.SubFolders
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            If LCase$(oSub.Name) <> "oldversions" Then
                PRJ_CollectByExt oSub.Path, True, sExt, colOut
            End If
            DoEvents
        Next oSub
        On Error GoTo 0
    End If
End Sub

' END: modSummary v9.19.9

