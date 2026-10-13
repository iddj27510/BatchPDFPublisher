Attribute VB_Name = "modInspect"
Option Explicit
' Module    : modInspect
' Purpose   : Read-only inspection of IDW sheet counts and PDF page counts
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.7 (Silent scan):
'   IDW_SheetCountViaInventor now brackets its open/close
'   cycle with ThisApplication.SilentOperation. Inventor
'   2026+ shows "Drawing updates are deferred" once per IDW
'   when a drawing with unresolved references is opened via
'   the full Inventor engine. The scan phase (frmScan) has no
'   SilentOperation set, so every file triggered the dialog.
'   SilentOperation is saved, set to True, and restored on
'   both normal and error exit.
' v9.19.6 (Documentation accuracy):
'   Notes section corrected: this module DOES use a line-continuation
'   character in PRJ_CheckPDFStatus (the BuildPath call). Prior wording
'   ("No line-continuation characters in this module") was inaccurate.
'   Requires section corrected: added PRJ_GetFileDate (modFileInfo),
'   used by PRJ_CheckPDFStatus for the stale-vs-OK comparison.
'   No functional changes.
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.19.5: Header sync to PRJ_VERSION. No functional changes.
' v9.19.4 (Header accuracy):
'   Header and inline comments corrected: PRJ_GetPDFPageCount is
'   a 2-tier lookup (cache -> Windows Property via Shell.Application),
'   not 3-tier. Prior header claimed "cache -> Windows Property ->
'   Acrobat PDDoc" but the Acrobat fallback was never implemented.
'   If a future revision adds the Acrobat PDDoc fallback, update
'   this header and Doc 06 / Doc 09.
' v9.19.3 (Thin wrapper cleanup):
'   Removed PDF_CreateShellApp — it was a one-line delegate to
'   SHELL_GetShellApp.
' v9.18.0: PRJ_GetIDWSheetCount fallback triggered on n < 0.
'          Err.Clear before CreateObject / ExtendedProperty calls.
'          PRJ_HasDrawing guard against paths without a dot.
'
' Requires  : CACHE_GetSheetCount, CACHE_GetPageCount,
'             CACHE_PutSheetCount, CACHE_PutPageCount (modCache).
'             PRJ_GetApprentice (modCOM).
'             PRJ_FSO (modCOM).
'             SHELL_GetShellApp (modShellApp).
'             PRJ_GetFileDate (modFileInfo).
'             PRJ_LogWarn (modLog).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-argument call
'             continuation are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private m_dFECache  As Object
Private m_sFEFolder As String

Public Function PRJ_GetIDWSheetCount(sIDW As String) As Long
    Dim nCached As Long
    nCached = CACHE_GetSheetCount(sIDW)
    If nCached > 0 Then
        PRJ_GetIDWSheetCount = nCached
        Exit Function
    End If
    Dim n As Long
    n = IDW_SheetCountViaApprentice(sIDW)
    If n < 0 Then
        PRJ_LogWarn "PRJ_GetIDWSheetCount", "Apprentice failed for " & sIDW & " - falling back to full open"
        n = IDW_SheetCountViaInventor(sIDW)
    End If
    If n > 0 Then CACHE_PutSheetCount sIDW, n
    PRJ_GetIDWSheetCount = n
End Function

Private Function IDW_SheetCountViaApprentice(sIDW As String) As Long
    IDW_SheetCountViaApprentice = -1
    Dim oDoc As Object
    On Error GoTo Fail
    Dim oApp As Object
    Set oApp = PRJ_GetApprentice()
    If oApp Is Nothing Then Exit Function
    Set oDoc = oApp.Open(sIDW)
    If oDoc Is Nothing Then Exit Function
    IDW_SheetCountViaApprentice = oDoc.Sheets.Count
    oDoc.Close
    Set oDoc = Nothing
    Exit Function
Fail:
    On Error Resume Next
    If Not oDoc Is Nothing Then oDoc.Close
    Set oDoc = Nothing
    On Error GoTo 0
    IDW_SheetCountViaApprentice = -1
End Function

Private Function IDW_SheetCountViaInventor(sIDW As String) As Long
    IDW_SheetCountViaInventor = -1
    Dim oDoc As Object
    Dim bWeOpened As Boolean
    Dim bWasSilent As Boolean
    Dim bRestore As Boolean

    ' v9.19.7: suppress the "Drawing updates are deferred" dialog
    ' that Inventor shows when a drawing is opened with unresolved
    ' references. Without SilentOperation, the dialog appears once
    ' per file and blocks the scan on folders with missing children.
    bWasSilent = False
    bRestore = False

    On Error Resume Next
    Err.Clear
    bWasSilent = ThisApplication.SilentOperation
    ThisApplication.SilentOperation = True
    bRestore = (Err.Number = 0)
    If Not bRestore Then Err.Clear
    On Error GoTo Fail

    On Error Resume Next
    Err.Clear
    Set oDoc = ThisApplication.Documents.ItemByName(sIDW)
    On Error GoTo Fail
    If oDoc Is Nothing Then
        Dim oOpts As NameValueMap
        Set oOpts = ThisApplication.TransientObjects.CreateNameValueMap
        oOpts.Add "SkipAllUnresolvedFiles", True
        Set oDoc = ThisApplication.Documents.OpenWithOptions(sIDW, oOpts, False)
        bWeOpened = True
    End If
    If oDoc Is Nothing Then GoTo Cleanup
    IDW_SheetCountViaInventor = oDoc.Sheets.Count
    If bWeOpened Then oDoc.Close False
    Set oDoc = Nothing
    GoTo Cleanup

Fail:
    On Error Resume Next
    If bWeOpened And Not oDoc Is Nothing Then oDoc.Close False
    Set oDoc = Nothing
    IDW_SheetCountViaInventor = -1

Cleanup:
    If bRestore Then
        On Error Resume Next
        ThisApplication.SilentOperation = bWasSilent
        On Error GoTo 0
    End If
End Function

Public Function PRJ_GetPDFPageCount(sPDF As String) As Long
    Dim n As Long
    n = CACHE_GetPageCount(sPDF)
    If n > 0 Then
        PRJ_GetPDFPageCount = n
        Exit Function
    End If
    n = PDF_PageCountViaShell(sPDF)
    If n > 0 Then
        CACHE_PutPageCount sPDF, n
        PRJ_GetPDFPageCount = n
        Exit Function
    End If
    PRJ_GetPDFPageCount = -1
End Function

Private Function PDF_PageCountViaShell(sPDF As String) As Long
    PDF_PageCountViaShell = -1
    Dim oShell As Object, oFolder As Object, oFile As Object
    Dim sFolder As String, sName As String
    sFolder = PRJ_FSO().GetParentFolderName(sPDF)
    sName = PRJ_FSO().GetFileName(sPDF)
    ' v9.19.3: inlined PDF_CreateShellApp — direct call to the
    ' Shell.Application factory isolated in modShellApp.
    Set oShell = SHELL_GetShellApp()
    If oShell Is Nothing Then Exit Function
    Set oFolder = PDF_GetShellFolder(oShell, sFolder)
    If oFolder Is Nothing Then Exit Function
    Set oFile = PDF_GetShellFile(oFolder, sName)
    If oFile Is Nothing Then Exit Function
    Dim v As Variant
    On Error Resume Next
    Err.Clear
    v = oFile.ExtendedProperty("System.Document.PageCount")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    Dim n As Long
    n = CLng(v)
    If n <= 0 Then Exit Function
    PDF_PageCountViaShell = n
End Function

Private Function PDF_GetShellFolder(ByVal oShell As Object, ByVal sPath As String) As Object
    On Error Resume Next
    Err.Clear
    Set PDF_GetShellFolder = oShell.Namespace(sPath)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function PDF_GetShellFile(ByVal oFolder As Object, ByVal sName As String) As Object
    On Error Resume Next
    Err.Clear
    Set PDF_GetShellFile = oFolder.ParseName(sName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function FILE_FolderCache(ByVal sFolder As String) As Object
    On Error Resume Next

    If m_dFECache Is Nothing Then
        Err.Clear
        Set m_dFECache = CreateObject("Scripting.Dictionary")
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Set FILE_FolderCache = Nothing
            Exit Function
        End If
        m_dFECache.CompareMode = vbTextCompare
    End If

    If StrComp(m_sFEFolder, sFolder, vbTextCompare) = 0 Then
        Set FILE_FolderCache = m_dFECache
        Exit Function
    End If

    m_dFECache.RemoveAll
    m_sFEFolder = sFolder

    If Left$(sFolder, 2) <> "\\" Then
        Set FILE_FolderCache = m_dFECache
        Exit Function
    End If

    Dim oFld As Object
    Dim oF As Object
    Set oFld = PRJ_FSO().GetFolder(sFolder)
    If Not oFld Is Nothing Then
        For Each oF In oFld.Files
            m_dFECache(oF.Name) = True
        Next
    End If

    Set FILE_FolderCache = m_dFECache
    On Error GoTo 0
End Function

Public Function FILE_ExistsCached(ByVal sPath As String) As Boolean
    On Error Resume Next

    Dim sFolder As String
    Dim sName   As String
    Dim oCache  As Object

    sFolder = PRJ_FSO().GetParentFolderName(sPath)
    sName = PRJ_FSO().GetFileName(sPath)

    Set oCache = FILE_FolderCache(sFolder)

    If oCache Is Nothing Or oCache.Count = 0 Then
        FILE_ExistsCached = PRJ_FSO().FileExists(sPath)
    Else
        FILE_ExistsCached = oCache.Exists(sName)
    End If
    On Error GoTo 0
End Function

Public Function PRJ_CheckPDFStatus(sIDW As String) As String
    Dim sPDF As String
    sPDF = PRJ_FSO().BuildPath(PRJ_FSO().GetParentFolderName(sIDW), _
                                PRJ_FSO().GetBaseName(sIDW) & ".pdf")
    If Not FILE_ExistsCached(sPDF) Then
        PRJ_CheckPDFStatus = "Missing"
        Exit Function
    End If
    Dim nPDF As Long
    Dim nIdw As Long
    nPDF = PRJ_GetPDFPageCount(sPDF)
    nIdw = PRJ_GetIDWSheetCount(sIDW)
    If nPDF > 0 And nIdw > 0 And nPDF <> nIdw Then
        PRJ_CheckPDFStatus = "Mismatch:" & nPDF & "/" & nIdw
        Exit Function
    End If
    If PRJ_GetFileDate(sIDW) > PRJ_GetFileDate(sPDF) Then
        PRJ_CheckPDFStatus = "Stale"
    Else
        PRJ_CheckPDFStatus = "OK"
    End If
End Function

Public Function PRJ_HasDrawing(sModel As String) As Boolean
    Dim sIDW As String
    Dim nDot As Long

    nDot = InStrRev(sModel, ".")
    If nDot = 0 Then
        PRJ_HasDrawing = False
        Exit Function
    End If

    sIDW = Left$(sModel, nDot) & "idw"
    PRJ_HasDrawing = FILE_ExistsCached(sIDW)
End Function

' END: modInspect v9.19.9

