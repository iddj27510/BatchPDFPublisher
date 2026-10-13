Attribute VB_Name = "modStatus"
Option Explicit
' Module    : modStatus
' Purpose   : Canonical status classifier for the report pipeline.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Fixed stale inline comment: validate pipeline was labelled
'   "(v9.19.4)" but the statuses were introduced in v9.19.6.
'   No functional changes.
'
' v9.19.6 (Validate status support):
'   Added ST_MISSING, ST_STALE, ST_MISMATCH, ST_OK to the TStatus
'   enum. These are emitted by PRJ_CheckPDFStatus (modInspect) and
'   consumed by SC_ValidateDrawings (modValidate). Prior to this
'   release SC_ValidateDrawings used a Select Case True block with
'   mixed string equality and Left$ prefix matching — a symptom of
'   the statuses not being in the central classifier. Now unified.
'
' v9.19.3 (Header standardization):
'   Header converted from Style A to Style B. Magic number lengths
'   (5, 7, 9, 10, 11, 15) replaced with Len() via Const declarations.
'
' Precedence rules — ORDER IS LOAD-BEARING:
'   1. "Skipped (Empty)" MUST be tested before "Skipped".
'   2. "Error" before "Overwritten"/"Created" — error messages can
'      contain the word "created" in the human text.
'   3. Prefix matching uses Left$ + Len(). Mirrors
'      modReportBuild.RP_LocalizeStatus which strips the same
'      characters via Mid(sRes, N+1).
'   4. "Mismatch" uses prefix match (matches "Mismatch:3/5").
'   5. "OK", "Missing", "Stale" use exact match.
'
' Requires  : (none)
' Notes     : No line-continuation characters in this module.
'             No decorative separator comment lines.

Public Enum TStatus
    ST_UNKNOWN = 0
    ST_CREATED = 1
    ST_OVERWRITTEN = 2
    ST_SKIPPED = 3
    ST_SKIPPED_EMPTY = 4
    ST_ERROR = 5
    ST_HAS_DRAWING = 6
    ST_NO_DRAWING = 7
    ' v9.19.6 — validate-pipeline statuses
    ST_MISSING = 8
    ST_STALE = 9
    ST_MISMATCH = 10
    ST_OK = 11
End Enum

' STS_Classify   parse a status string into a TStatus value.
'   Empty input -> ST_UNKNOWN. Unknown prefix -> ST_UNKNOWN.
'   Case-sensitive on the prefix; statuses are emitted by our own
'   code so casing is stable.
'
'   Order of checks below is not arbitrary — see the file header
'   for the precedence rules. Reordering is a bug.
' Params  : sRes — String, canonical status emitted by modPDFExport,
'           modInspect.PRJ_CheckPDFStatus, or the audit collector
' Returns : TStatus — one of the 12 enum members
' Throws  : (none)
Public Function STS_Classify(ByVal sRes As String) As TStatus
    Const S_HAS_DRAWING   As String = "HasDrawing"
    Const S_NO_DRAWING    As String = "NoDrawing"
    Const S_SKIPPED_EMPTY As String = "Skipped (Empty)"
    Const S_SKIPPED       As String = "Skipped"
    Const S_ERROR         As String = "Error"
    Const S_OVERWRITTEN   As String = "Overwritten"
    Const S_CREATED       As String = "Created"
    Const S_MISSING       As String = "Missing"
    Const S_STALE         As String = "Stale"
    Const S_MISMATCH      As String = "Mismatch"
    Const S_OK            As String = "OK"

    STS_Classify = ST_UNKNOWN

    Dim nLen As Long
    nLen = Len(sRes)
    If nLen = 0 Then Exit Function

    ' -- Audit-mode rows (exact standalone words) -----------------
    If nLen >= Len(S_HAS_DRAWING) Then
        If Left$(sRes, Len(S_HAS_DRAWING)) = S_HAS_DRAWING Then
            STS_Classify = ST_HAS_DRAWING
            Exit Function
        End If
    End If
    If nLen >= Len(S_NO_DRAWING) Then
        If Left$(sRes, Len(S_NO_DRAWING)) = S_NO_DRAWING Then
            STS_Classify = ST_NO_DRAWING
            Exit Function
        End If
    End If

    ' -- PDF export pipeline --------------------------------------
    If nLen >= Len(S_SKIPPED_EMPTY) Then
        If Left$(sRes, Len(S_SKIPPED_EMPTY)) = S_SKIPPED_EMPTY Then
            STS_Classify = ST_SKIPPED_EMPTY
            Exit Function
        End If
    End If
    If nLen >= Len(S_SKIPPED) Then
        If Left$(sRes, Len(S_SKIPPED)) = S_SKIPPED Then
            STS_Classify = ST_SKIPPED
            Exit Function
        End If
    End If
    If nLen >= Len(S_ERROR) Then
        If Left$(sRes, Len(S_ERROR)) = S_ERROR Then
            STS_Classify = ST_ERROR
            Exit Function
        End If
    End If
    If nLen >= Len(S_OVERWRITTEN) Then
        If Left$(sRes, Len(S_OVERWRITTEN)) = S_OVERWRITTEN Then
            STS_Classify = ST_OVERWRITTEN
            Exit Function
        End If
    End If
    If nLen >= Len(S_CREATED) Then
        If Left$(sRes, Len(S_CREATED)) = S_CREATED Then
            STS_Classify = ST_CREATED
            Exit Function
        End If
    End If

    ' -- Validate pipeline (v9.19.6) ------------------------------
    ' "Mismatch" is a prefix match because the actual string is
    ' "Mismatch:3/5" (pages/page mismatch).
    If nLen >= Len(S_MISMATCH) Then
        If Left$(sRes, Len(S_MISMATCH)) = S_MISMATCH Then
            STS_Classify = ST_MISMATCH
            Exit Function
        End If
    End If
    ' Exact matches for the three remaining validate statuses.
    If sRes = S_MISSING Then
        STS_Classify = ST_MISSING
        Exit Function
    End If
    If sRes = S_STALE Then
        STS_Classify = ST_STALE
        Exit Function
    End If
    If sRes = S_OK Then
        STS_Classify = ST_OK
        Exit Function
    End If
End Function

' END: modStatus v9.19.9
