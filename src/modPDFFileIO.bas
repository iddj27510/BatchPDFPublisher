Attribute VB_Name = "modPDFFileIO"
Option Explicit
' Module    : modPDFFileIO
' Purpose   : Temp file lifecycle and atomic replace for PDF pipeline
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + header completion):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Added missing Reference field for consistency with the rest
'   of the project.
'   Fixed Requires section: PDF_* constants live in modConstants,
'   not modState. The prior wording ("(modState)") was incorrect.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in PDF_BuildTempPath, PDF_FindAddIn and
'   PDF_CommitReplace). No functional changes.
'
' v9.18.0 (§9.10 API versioning):
'   Added PDF_FindAddIn (primary GUID lookup, fallback enumeration
'   by DisplayName containing "PDF"). PDF_ExportToTemp now uses
'   PDF_FindAddIn instead of direct ApplicationAddIns.ItemById call.
' v9.18.0 (§8.11 File system paths):
'   PDF_BuildTempPath uses PRJ_PathCombine.
' v9.18.0: Err.Clear before each API call.
'          PDF_ExportToTemp EH deletes partial temp file if add-in failed.
' v9.13.9: removed local PDF_TEMP_* constants - now in modState.
'
' Requires  : TEMP_FILE_PREFIX, PDF_TEMP_SEG_MAX, PDF_TEMP_RAND_MAX,
'             PDF_GUID, PDF_VECTOR_RESOLUTION, PDF_ALL_COLOR_AS_BLACK,
'             PDF_PDF_STANDARD (modConstants).
'             g_nTempCounter (modState).
'             RETRY_DeleteFile, RETRY_MoveFile (modRetry).
'             PRJ_FSO (modCOM).
'             PRJ_SandboxPath, PRJ_PathCombine (modSandbox).
'             PRJ_LogWarn / PRJ_LogError / PRJ_LogInfo / PRJ_LogDebug (modLog).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-argument call
'             continuation are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Public Function PDF_BuildTempPath(sPDF As String) As String
    Dim sSafePath As String
    Dim sPid As String
    Dim sUniq As String

    sSafePath = LCase$(sPDF)
    sSafePath = Replace(sSafePath, "\", "_")
    sSafePath = Replace(sSafePath, ":", "")
    sSafePath = Replace(sSafePath, " ", "_")
    sSafePath = Replace(sSafePath, "/", "_")

    If Len(sSafePath) > PDF_TEMP_SEG_MAX Then sSafePath = Right$(sSafePath, PDF_TEMP_SEG_MAX)

    sPid = Format$(GetCurrentProcessId(), "00000")

    g_nTempCounter = g_nTempCounter + 1
    sUniq = Format(Timer, "0.000000") & "_" & Format(g_nTempCounter, "0000") & "_" & Format(Int(Rnd() * PDF_TEMP_RAND_MAX), "0000")
    sUniq = Replace(sUniq, ".", "")
    sUniq = Replace(sUniq, ",", "")

    PDF_BuildTempPath = PRJ_PathCombine(PRJ_SandboxPath(), _
                        TEMP_FILE_PREFIX & sPid & "_" & sUniq & "_" & sSafePath & ".pdf")
End Function

Private Function PDF_FindAddIn() As TranslatorAddIn
    ' §9.10: Primary lookup by known GUID. Fallback: enumerate
    ' ApplicationAddIns and match by DisplayName containing "PDF".
    ' Handles the case where Autodesk changes the PDF Translator
    ' AddIn GUID in a future Inventor version.
    Set PDF_FindAddIn = Nothing

    On Error Resume Next
    Err.Clear
    Set PDF_FindAddIn = ThisApplication.ApplicationAddIns.ItemById(PDF_GUID)
    If Err.Number <> 0 Then
        Err.Clear
        Set PDF_FindAddIn = Nothing
    End If
    On Error GoTo 0
    If Not PDF_FindAddIn Is Nothing Then Exit Function

    Dim oAddIn As Object
    On Error Resume Next
    For Each oAddIn In ThisApplication.ApplicationAddIns
        If InStr(1, oAddIn.DisplayName, "PDF", vbTextCompare) > 0 Then
            Set PDF_FindAddIn = oAddIn
            Exit For
        End If
    Next
    On Error GoTo 0

    If Not PDF_FindAddIn Is Nothing Then
        PRJ_LogWarn "PDF_FindAddIn", _
            "PDF Translator found via enumeration, not GUID. GUID may have changed: " & PDF_GUID
    End If
End Function

Private Sub PDF_ApplyAddInOption(ByRef oOpts As Object, ByVal sKey As String, ByVal vValue As Variant)
    On Error Resume Next
    Err.Clear
    oOpts.Value(sKey) = vValue
    If Err.Number <> 0 Then
        Err.Clear
        oOpts.Add sKey, vValue
        If Err.Number <> 0 Then Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Function PDF_ExportToTemp(oDoc As DrawingDocument, sTemp As String) As String
    PDF_ExportToTemp = ""

    Dim oAddIn As TranslatorAddIn
    Set oAddIn = PDF_FindAddIn()
    If oAddIn Is Nothing Then
        PDF_ExportToTemp = "PDF Translator AddIn not found"
        Exit Function
    End If

    Dim oCtx As TranslationContext
    Set oCtx = ThisApplication.TransientObjects.CreateTranslationContext
    oCtx.Type = kFileBrowseIOMechanism

    Dim oOpts As NameValueMap
    Set oOpts = ThisApplication.TransientObjects.CreateNameValueMap

    Dim oData As DataMedium
    Set oData = ThisApplication.TransientObjects.CreateDataMedium

    If oAddIn.HasSaveCopyAsOptions(oDoc, oCtx, oOpts) Then
        PDF_ApplyAddInOption oOpts, "Publish_All_Sheets", True
        PDF_ApplyAddInOption oOpts, "Launch_Viewer", False
        PDF_ApplyAddInOption oOpts, "Vector_Resolution", PDF_VECTOR_RESOLUTION
        PDF_ApplyAddInOption oOpts, "All_Color_AS_Black", PDF_ALL_COLOR_AS_BLACK
        If Len(PDF_PDF_STANDARD) > 0 Then
            PDF_ApplyAddInOption oOpts, "PDF_Standard", PDF_PDF_STANDARD
        End If
    End If

    On Error Resume Next
    If PRJ_FSO().FileExists(sTemp) Then PRJ_FSO().DeleteFile sTemp, True
    On Error GoTo 0

    On Error GoTo EH
    oData.FileName = sTemp
    oAddIn.SaveCopyAs oDoc, oCtx, oOpts, oData
    On Error GoTo 0

    If Not PRJ_FSO().FileExists(sTemp) Then
        PDF_ExportToTemp = "PDF export produced no file"
    End If
    Exit Function

EH:
    On Error Resume Next
    If PRJ_FSO().FileExists(sTemp) Then PRJ_FSO().DeleteFile sTemp, True
    On Error GoTo 0
    PDF_ExportToTemp = "Export failed: " & Err.Description
End Function

Public Function PDF_ValidateHeader(sPath As String) As Boolean
    PDF_ValidateHeader = False
    Dim ff As Integer
    On Error GoTo Fail

    If Len(sPath) = 0 Then Exit Function
    If Not PRJ_FSO().FileExists(sPath) Then Exit Function
    If PRJ_FSO().GetFile(sPath).Size < 5 Then Exit Function

    ff = FreeFile
    Open sPath For Binary Access Read As #ff
    Dim b(4) As Byte
    Get #ff, 1, b
    Close #ff

    PDF_ValidateHeader = (b(0) = &H25 And b(1) = &H50 And b(2) = &H44 And b(3) = &H46 And b(4) = &H2D)
    Exit Function

Fail:
    On Error Resume Next
    If ff > 0 Then Close #ff
    On Error GoTo 0
    PDF_ValidateHeader = False
End Function

Public Function PDF_SafeReplace(sTemp As String, sPDF As String) As String
    Dim sBak       As String
    Dim bHadOldPdf As Boolean

    PDF_SafeReplace = ""
    sBak = sPDF & ".bak"
    bHadOldPdf = False

    If PRJ_FSO().FileExists(sPDF) Then
        If PRJ_FSO().FileExists(sBak) Then
            If Not RETRY_DeleteFile(sBak) Then
                PDF_SafeReplace = "Cannot remove stale backup: " & sBak
                Exit Function
            End If
        End If
        If Not RETRY_MoveFile(sPDF, sBak) Then
            PDF_SafeReplace = "Cannot back up existing PDF: " & sPDF
            Exit Function
        End If
        bHadOldPdf = True
    End If

    If Not RETRY_MoveFile(sTemp, sPDF) Then
        If bHadOldPdf Then
            If PRJ_FSO().FileExists(sPDF) Then RETRY_DeleteFile sPDF
            RETRY_MoveFile sBak, sPDF
        End If
        PDF_SafeReplace = "Cannot write PDF: " & sPDF
        Exit Function
    End If
End Function

Public Sub PDF_CommitReplace(sPDF As String)
    Dim sBak As String
    sBak = sPDF & ".bak"
    If Not PRJ_FSO().FileExists(sBak) Then Exit Sub
    If Not RETRY_DeleteFile(sBak) Then
        PRJ_LogWarn "PDF_CommitReplace", "Backup not removed after successful replace: " & sBak & " (will retry on next export)"
    Else
        PRJ_LogInfo "AUDIT", "Deleted backup: " & sBak
    End If
End Sub

Public Sub PDF_RollbackReplace(sPDF As String)
    Dim sBak As String
    sBak = sPDF & ".bak"
    If Not PRJ_FSO().FileExists(sBak) Then Exit Sub

    If PRJ_FSO().FileExists(sPDF) Then
        If Not RETRY_DeleteFile(sPDF) Then
            PRJ_LogError "PDF_RollbackReplace", "Cannot delete bad PDF " & sPDF & "; rollback aborted, backup retained at " & sBak
            Exit Sub
        End If
    End If

    If Not RETRY_MoveFile(sBak, sPDF) Then
        PRJ_LogError "PDF_RollbackReplace", "Rollback failed: could not restore " & sBak & " to " & sPDF & "; manual recovery needed"
    End If
End Sub

' END: modPDFFileIO v9.19.9

