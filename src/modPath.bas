Attribute VB_Name = "modPath"
Option Explicit
' Module    : modPath
' Purpose   : Path sanitization pipeline
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + header language):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   v9.19.3 changelog entry translated from Serbian to English
'   (the rest of the module was already in English).
'   Fixed encoding artifact in the v9.19.3 note: '?' was a broken
'   arrow; replaced with '->'. No functional changes.
'
' v9.19.3 (Root clamp fix):
'   SN_ResolveParentSegments: the Else branch (no backslash before
'   "\..\") now keeps the prefix instead of discarding it. Prior to
'   the fix, "C:\Foo\..\..\Bar" returned "Bar" (drive lost).
'   After the fix it returns "C:\Bar" (Windows semantics: parent
'   traversal cannot go above the root).
'   Bug covered by test Test_SN_SanitizePath -> "root clamp".
' v9.18.0: Moved bWasUNC / sShare extraction to AFTER SN_StripOuterQuotes.
'          Prior order (UNC detect before quote strip) missed the case
'          of a quoted UNC path with parent segments, e.g.
'          "\\srv\share\..\..\windows" - the leading char was ", so
'          bWasUNC was False and the UNC guard was skipped. Result:
'          share boundary was escapable.
'          Trim -> Trim$ (avoid Variant wrap).
' v9.16.8 (7.6b): XML doc annotations; inline WHY in SN_ResolveParentSegments.
' v9.15.3 (Faza 6C test): SN_SanitizePath verifies that the UNC share
'          (\\server\share) is preserved after resolution, not just the
'          \\ prefix.
' Notes     : v9.13.2 (ARCH-08): extracted from modUtils to satisfy SRP.
'             Order matters: quotes -> illegal chars -> duplicate slashes ->
'             parent segments -> trailing slash -> length check.
'             Changing the order breaks UNC handling (phase 4 after phase 3
'             restores the \\ prefix if the input was a UNC path) and
'             parent resolution (phase 5 needs a single backslash separator).
'             No line-continuation characters anywhere in this module.

'@Description "Sanitize a user-provided path for FSO / FileSystem use."
'@Param     "sIn" "Raw path string. May include outer quotes, illegal chars,
'                  mixed separators, parent segments (..), or trailing slash."
'@Returns   "Sanitized path, or empty string if invalid."
'@Remarks   "Returns empty string in three cases:
'              1. Input is empty after Trim.
'              2. Input was UNC but parent resolution escaped the share
'                 boundary (e.g. \\srv\share\..\..\Windows -> invalid).
'              3. Sanitized path exceeds MAX_PATH_LEN (260).
'            Callers should treat empty return as 'invalid input' and
'            refuse to proceed."
Public Function SN_SanitizePath(sIn As String) As String
    Dim s As String
    Dim bWasUNC As Boolean
    Dim sShare As String

    s = Trim$(sIn)
    If Len(s) = 0 Then
        SN_SanitizePath = ""
        Exit Function
    End If

    ' Strip outer quotes first, so that a quoted UNC path like
    ' "\"\\srv\share\..\..\win\"" is correctly classified as UNC.
    s = SN_StripOuterQuotes(s)

    bWasUNC = (Left$(s, 2) = "\\")
    If bWasUNC Then sShare = SN_GetUNCShare(s)

    s = SN_StripIllegalChars(s)
    s = SN_CollapseDblBackslashes(s)
    s = SN_ResolveParentSegments(s)
    s = SN_StripTrailingBackslash(s)

    ' UNC guard: parent traversal must not escape the share boundary.
    ' Pre-9.15.3 check only verified the \\ prefix survived; that allowed
    ' \\server\share\..\..\Windows -> \\Windows to pass. Now we verify
    ' the actual share (\\server\share) is preserved.
    If bWasUNC Then
        If Len(sShare) = 0 Then
            SN_SanitizePath = ""
            Exit Function
        End If
        If Len(s) < Len(sShare) Then
            SN_SanitizePath = ""
            Exit Function
        End If
        If LCase$(Left$(s, Len(sShare))) <> LCase$(sShare) Then
            SN_SanitizePath = ""
            Exit Function
        End If
    End If

    If Len(s) > MAX_PATH_LEN Then
        SN_SanitizePath = ""
        Exit Function
    End If

    SN_SanitizePath = s
End Function

' SN_GetUNCShare   extract "\\server\share" from a UNC path.
'   Returns "" if input is not a valid UNC share (no server or no share).
Private Function SN_GetUNCShare(ByVal s As String) As String
    Dim p1 As Long
    Dim p2 As Long

    SN_GetUNCShare = ""
    If Left$(s, 2) <> "\\" Then Exit Function

    p1 = InStr(3, s, "\")
    If p1 = 0 Then Exit Function

    p2 = InStr(p1 + 1, s, "\")
    If p2 = 0 Then
        SN_GetUNCShare = s
    Else
        SN_GetUNCShare = Left$(s, p2 - 1)
    End If
End Function

' SN_StripOuterQuotes   remove a matched pair of double quotes around the
'   whole path. Users often paste from Explorer's "Copy as path" which
'   wraps the result in quotes; without stripping, the quotes become part
'   of the path and FSO rejects it.
Private Function SN_StripOuterQuotes(ByVal s As String) As String
    SN_StripOuterQuotes = s
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) = """" And Right$(s, 1) = """" Then
        SN_StripOuterQuotes = Trim$(Mid$(s, 2, Len(s) - 2))
    End If
End Function

' SN_StripIllegalChars   drop control chars (0..31) and the reserved
'   Windows filename characters " * ? < > |.
Private Function SN_StripIllegalChars(ByVal s As String) As String
    Dim i     As Long
    Dim ch    As String
    Dim nCode As Long
    Dim sOut  As String

    sOut = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        nCode = AscW(ch)
        If nCode < 0 Then nCode = nCode + 65536
        Select Case nCode
            Case 0 To 31, 34, 42, 63, 60, 62, 124
                ' skip
            Case Else
                sOut = sOut & ch
        End Select
    Next i
    SN_StripIllegalChars = sOut
End Function

' SN_CollapseDblBackslashes   collapse runs of \\ into a single \.
'   UNC paths (\\server\share) start with a legal double backslash; the
'   collapse scan starts at position 3 for those so the UNC prefix is
'   preserved. Any later \\ in the path is a user error or paste artifact.
Private Function SN_CollapseDblBackslashes(ByVal s As String) As String
    Dim nStart As Long
    Dim nDbl   As Long

    nStart = 1
    If Left$(s, 2) = "\\" Then nStart = 3
    Do While InStr(nStart, s, "\\") > 0
        nDbl = InStr(nStart, s, "\\")
        s = Left$(s, nDbl - 1) & Mid$(s, nDbl + 1)
    Loop
    SN_CollapseDblBackslashes = s
End Function

' SN_ResolveParentSegments   walk the string resolving "\..\" pairs by
'   deleting the preceding segment.
'   Capped at PATH_MAX_PARENT_ITER (20) iterations to guarantee termination
'   on pathological inputs (e.g. "\..\..\..\..\..." repeated many times).
'   Without the cap, a crafted input with hundreds of parent segments
'   could loop for a long time or (theoretically) hang on an edge case.
'
'   When no backslash precedes "\..\" (we're at drive root like "C:" or
'   root-relative start), the traversal cannot go up further — Windows
'   clamps at the root. In that case we keep what's before (drive prefix
'   like "C:", or "" for root-relative paths) and continue from the rest
'   after the "\..\". This matches the v9.19.3 fix: pre-fix behavior
'   dropped the drive prefix ("C:\Foo\..\..\Bar" -> "Bar").
'
'   NOTE: does not protect the UNC share boundary itself - that check
'   lives in SN_SanitizePath (see UNC guard there).
Private Function SN_ResolveParentSegments(ByVal s As String) As String
    Dim nIter   As Long
    Dim nPos    As Long
    Dim sBefore As String
    Dim nSlash  As Long

    nIter = 0
    Do While InStr(s, "\..\") > 0 And nIter < PATH_MAX_PARENT_ITER
        nIter = nIter + 1
        nPos = InStr(s, "\..\")
        sBefore = Left$(s, nPos - 1)
        nSlash = InStrRev(sBefore, "\")
        If nSlash > 0 Then
            ' Normal case: "C:\Foo\..\Bar" -> "C:\Bar"
            s = Left$(s, nSlash - 1) & Mid$(s, nPos + 3)
        Else
            ' At root: "C:\..\Bar" or "\..\Bar"
            ' Cannot go up further — keep the prefix (drive "C:" or "")
            ' and continue after the \..\. This is the v9.19.3 fix.
            s = sBefore & Mid$(s, nPos + 3)
        End If
    Loop
    SN_ResolveParentSegments = s
End Function

' SN_StripTrailingBackslash   remove a trailing \ unless it is the drive
'   root (C:\ has length 3; we stop at Len = 3 so C:\ survives intact).
Private Function SN_StripTrailingBackslash(ByVal s As String) As String
    Do While Len(s) > 3 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    SN_StripTrailingBackslash = s
End Function

' END: modPath v9.19.9
