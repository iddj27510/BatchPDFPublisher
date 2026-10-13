Attribute VB_Name = "modQA"
Option Explicit
' Module    : modQA
' Purpose   : Quality Assurance - IDW and PDF content validation
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + Requires fix):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Fixed Requires section: QA_STAT_* and QA_PDF_MIN_BYTES live in
'   modConstants, not modState. The prior wording ("all in
'   modState") was incorrect — verified against the actual
'   declarations in modConstants.
'   Added the standard "No line-continuation characters" note
'   for consistency with the rest of the project.
'   No functional changes.
'
' v9.19.3 (Header sync):
'   Version bumped to match PRJ_VERSION. No code changes.
' v9.18.0:
'   QA_AnalyzePDF re-indented (was column-0). Err.Clear before
'   PRJ_GetFileSize and oSheet.Name (stale-Err hygiene).
' v9.14.1: Initial.
'
' Notes     : Called from modPDFExport pipeline.
'             Returns semicolon-separated warnings.
'             Empty string = QA passed.
'             No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.
'
' Requires  : g_bAnalyzeIDW, g_bAnalyzePDF (modState).
'             QA_STAT_IDW_EMPTY_VIEW, QA_STAT_IDW_NO_TITLE,
'             QA_STAT_PDF_TOO_SMALL, QA_PDF_MIN_BYTES (modConstants).
'             PRJ_GetFileSize / PRJ_FormatFileSize (modFileInfo).
'             PRJ_LogWarn (modLog).

Public Function QA_AnalyzeIDW(ByVal oDoc As Object) As String
    QA_AnalyzeIDW = ""
    If Not g_bAnalyzeIDW Then Exit Function
    If oDoc Is Nothing Then Exit Function

    Dim sWarnings As String
    Dim oSheet    As Object
    Dim oView     As Object
    Dim nViews    As Long
    Dim bAnyTitle As Boolean
    Dim bSheetErr As Boolean

    sWarnings = ""

    On Error GoTo EH

    ' 1. Empty views per sheet
    For Each oSheet In oDoc.Sheets
        nViews = 0
        bSheetErr = False

        On Error Resume Next
        Err.Clear
        For Each oView In oSheet.DrawingViews
            nViews = nViews + 1
        Next
        If Err.Number <> 0 Then
            PRJ_LogWarn "QA_AnalyzeIDW", "View iteration failed for [" & oSheet.Name & "]: " & Err.Description
            Err.Clear
            bSheetErr = True
        End If
        On Error GoTo EH

        If Not bSheetErr And nViews = 0 Then
            sWarnings = sWarnings & QA_STAT_IDW_EMPTY_VIEW & "[" & oSheet.Name & "];"
        End If
    Next

    ' 2. Title block presence - any sheet
    bAnyTitle = False
    For Each oSheet In oDoc.Sheets
        On Error Resume Next
        Err.Clear
        If Not oSheet.TitleBlock Is Nothing Then
            bAnyTitle = True
        End If
        If Err.Number <> 0 Then
            PRJ_LogWarn "QA_AnalyzeIDW", "TitleBlock check failed for [" & oSheet.Name & "]: " & Err.Description
            Err.Clear
        End If
        On Error GoTo EH
        If bAnyTitle Then Exit For
    Next

    If Not bAnyTitle Then
        sWarnings = sWarnings & QA_STAT_IDW_NO_TITLE & ";"
    End If

    QA_AnalyzeIDW = sWarnings
    Exit Function

EH:
    PRJ_LogWarn "QA_AnalyzeIDW", "Unexpected error: " & Err.Description
    QA_AnalyzeIDW = sWarnings
End Function

Public Function QA_AnalyzePDF(ByVal sPDFPath As String) As String
    Dim sWarnings As String
    Dim dSize     As Double

    QA_AnalyzePDF = ""
    If Not g_bAnalyzePDF Then Exit Function

    sWarnings = ""

    On Error Resume Next
    Err.Clear
    dSize = PRJ_GetFileSize(sPDFPath)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        QA_AnalyzePDF = sWarnings
        Exit Function
    End If
    On Error GoTo 0

    ' Size sanity: PDF smaller than QA_PDF_MIN_BYTES is suspicious.
    If dSize >= 0 And dSize < QA_PDF_MIN_BYTES Then
        sWarnings = sWarnings & QA_STAT_PDF_TOO_SMALL & "(" & PRJ_FormatFileSize(dSize) & ");"
    End If

    QA_AnalyzePDF = sWarnings
End Function

' END: modQA v9.19.9
