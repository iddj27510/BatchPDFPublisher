Attribute VB_Name = "modCache"
Option Explicit
' Module    : modCache
' Purpose   : Page count and sheet count caches with stale-entry
'             detection and FIFO eviction. Keys are file paths; values
'             are "stamp|count" strings where stamp is the file's
'             last-modified time. Cache hit requires stamp match.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + header standardization):
'   Header converted from Style A (free-form) to Style B
'   (Module/Purpose/Reference/Version). Added missing Requires
'   and Notes sections.
'   Changelog corrected: the v9.1 note claimed CACHE_MAX_ENTRIES /
'   CACHE_EVICT_BATCH were declared in modState, but they live in
'   modConstants. Corrected inline.
'   No functional changes.
'
' v9.19.6: CACHE_EvictBatch now reads dCache.Count after eviction
'          instead of computing (nCount - nEvict). The prior log line
'          reported a stale count; the new line reports the actual
'          remaining entries.
'          Header sync to PRJ_VERSION.
' v9.19.6: Format -> Format$ in CACHE_GetPageCount /
'          CACHE_GetSheetCount / CACHE_PutPageCount / CACHE_PutSheetCount
'          (project §3.19 convention — avoid Variant wrap on hot path).
'          Header sync to PRJ_VERSION.
' v9.18.0: ByVal dCache on CACHE_EvictIfNeeded / CACHE_EvictBatch
'          (Object is already a reference; ByRef adds nothing).
'          Header sync to PRJ_VERSION.
' v9.13.7 (RULE S.2): removed all decorative separator comment lines.
' v9.13.1: Header sync to PRJ_VERSION.
'          CACHE_Init: CompareMode = vbTextCompare. Windows paths are
'          case-insensitive, so without this C:\X.PDF and c:\x.pdf would
'          be two cache entries (duplicate lookups + inflated Count that
'          triggers premature eviction).
'          CACHE_Init: OERN on CreateObject now checks Err.Number and logs
'          failure instead of silently leaving the cache Nothing.
'          CACHE_EvictIfNeeded: split into orchestrator + CACHE_EvictBatch
'          for readability.
' v9.1 - OmniScope ULTRA fixes applied:
'   - Cache eviction implemented using CACHE_MAX_ENTRIES and
'     CACHE_EVICT_BATCH constants declared in modConstants      (A.33)
'   - On Error Resume Next scoped to minimum per operation      (A.15)
'   - CACHE_Clear uses RemoveAll instead of recreating objects  (A.102)
'   - Unreachable On Error GoTo 0 moved to correct position     (A.2)
' v9.1 changes:
'   - CACHE_PutPageCount / CACHE_PutSheetCount: split blanket OERN into
'     per-operation guards; abort write if stamp could not be obtained (CACHE-01)
'   - CACHE_EvictIfNeeded: added IsArray guard on Keys() result and
'     explicit nCount <= 0 exit before eviction loop                  (CACHE-02)
'
' Requires  : CACHE_MAX_ENTRIES, CACHE_EVICT_BATCH (modConstants).
'             g_dPageCache, g_dSheetCache (modState).
'             PRJ_GetFileDate (modFileInfo).
'             PRJ_LogError, PRJ_LogInfo (modLog).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

' Init / Clear

Public Sub CACHE_Init()
    ' Create both dictionaries if not already present.
    ' Called once from STATE_Init; safe to call again if needed.
    '
    ' CompareMode = vbTextCompare (case-insensitive) is critical for
    ' correctness: Windows file paths are case-insensitive, and the same
    ' file can arrive as C:\Foo\Bar.pdf or c:\foo\bar.pdf depending on
    ' how the user typed it, how Explorer returned it, or how the
    ' Inventor API normalized it. Binary compare (default) would treat
    ' these as distinct keys.
    '
    ' CompareMode can only be set while the Dictionary is empty, which is
    ' guaranteed here because we just created it.
    If g_dPageCache Is Nothing Then
        On Error Resume Next
        Err.Clear
        Set g_dPageCache = CreateObject("Scripting.Dictionary")
        If Err.Number <> 0 Or g_dPageCache Is Nothing Then
            Err.Clear
            On Error GoTo 0
            PRJ_LogError "modCache.CACHE_Init", "Cannot create g_dPageCache Dictionary"
        Else
            On Error Resume Next
            Err.Clear
            g_dPageCache.CompareMode = vbTextCompare
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    End If
    If g_dSheetCache Is Nothing Then
        On Error Resume Next
        Err.Clear
        Set g_dSheetCache = CreateObject("Scripting.Dictionary")
        If Err.Number <> 0 Or g_dSheetCache Is Nothing Then
            Err.Clear
            On Error GoTo 0
            PRJ_LogError "modCache.CACHE_Init", "Cannot create g_dSheetCache Dictionary"
        Else
            On Error Resume Next
            Err.Clear
            g_dSheetCache.CompareMode = vbTextCompare
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    End If
End Sub

' Invalidation

Public Sub CACHE_InvalidatePDF(ByVal sPDF As String)
    On Error Resume Next
    If Not g_dPageCache Is Nothing Then
        If g_dPageCache.Exists(sPDF) Then g_dPageCache.Remove sPDF
    End If
    On Error GoTo 0
End Sub

' Getters - returns -1 on miss, stale entry, or error

Public Function CACHE_GetPageCount(ByVal sPDF As String) As Long
    CACHE_GetPageCount = -1

    If g_dPageCache Is Nothing Then Exit Function

    Dim bExists As Boolean
    On Error Resume Next
    bExists = g_dPageCache.Exists(sPDF)
    On Error GoTo 0
    If Not bExists Then Exit Function

    Dim sEntry As String
    On Error Resume Next
    sEntry = CStr(g_dPageCache(sPDF))
    On Error GoTo 0

    Dim nPipe As Long
    nPipe = InStr(sEntry, "|")
    If nPipe = 0 Then Exit Function

    Dim sStamp As String
    Dim sNow   As String
    sStamp = Left(sEntry, nPipe - 1)

    On Error Resume Next
    sNow = Format$(PRJ_GetFileDate(sPDF), "yyyymmddhhnnss")
    On Error GoTo 0

    If sNow <> sStamp Then
        ' Stale entry - file modified since caching; remove and miss
        On Error Resume Next
        g_dPageCache.Remove sPDF
        On Error GoTo 0
        Exit Function
    End If

    On Error Resume Next
    CACHE_GetPageCount = CLng(Mid(sEntry, nPipe + 1))
    On Error GoTo 0
End Function

Public Function CACHE_GetSheetCount(ByVal sIDW As String) As Long
    CACHE_GetSheetCount = -1

    If g_dSheetCache Is Nothing Then Exit Function

    Dim bExists As Boolean
    On Error Resume Next
    bExists = g_dSheetCache.Exists(sIDW)
    On Error GoTo 0
    If Not bExists Then Exit Function

    Dim sEntry As String
    On Error Resume Next
    sEntry = CStr(g_dSheetCache(sIDW))
    On Error GoTo 0

    Dim nPipe As Long
    nPipe = InStr(sEntry, "|")
    If nPipe = 0 Then Exit Function

    Dim sStamp As String
    Dim sNow   As String
    sStamp = Left(sEntry, nPipe - 1)

    On Error Resume Next
    sNow = Format$(PRJ_GetFileDate(sIDW), "yyyymmddhhnnss")
    On Error GoTo 0

    If sNow <> sStamp Then
        ' Stale entry - file modified since caching; remove and miss
        On Error Resume Next
        g_dSheetCache.Remove sIDW
        On Error GoTo 0
        Exit Function
    End If

    On Error Resume Next
    CACHE_GetSheetCount = CLng(Mid(sEntry, nPipe + 1))
    On Error GoTo 0
End Function

' Putters - write to cache, enforce eviction limit after write
' FIX CACHE-01: blanket OERN split into per-operation guards.
'               If the file stamp cannot be obtained, the write is aborted
'               entirely rather than storing an empty-stamp entry that the
'               next Get would immediately treat as stale and purge.

Public Sub CACHE_PutPageCount(ByVal sPDF As String, ByVal nPages As Long)
    If g_dPageCache Is Nothing Then CACHE_Init
    If g_dPageCache Is Nothing Then Exit Sub

    Dim dDate As Date
    dDate = PRJ_GetFileDate(sPDF)
    If dDate = 0 Then Exit Sub

    Dim sStamp As String
    sStamp = Format$(dDate, "yyyymmddhhnnss")

    On Error Resume Next
    g_dPageCache(sPDF) = sStamp & "|" & CStr(nPages)
    On Error GoTo 0

    CACHE_EvictIfNeeded g_dPageCache
End Sub

Public Sub CACHE_PutSheetCount(ByVal sIDW As String, ByVal nSheets As Long)
    If g_dSheetCache Is Nothing Then CACHE_Init
    If g_dSheetCache Is Nothing Then Exit Sub

    Dim dDate As Date
    dDate = PRJ_GetFileDate(sIDW)
    If dDate = 0 Then Exit Sub

    Dim sStamp As String
    sStamp = Format$(dDate, "yyyymmddhhnnss")

    On Error Resume Next
    g_dSheetCache(sIDW) = sStamp & "|" & CStr(nSheets)
    On Error GoTo 0

    CACHE_EvictIfNeeded g_dSheetCache
End Sub

' CACHE_EvictIfNeeded - FIX A.33
' Original bug: CACHE_MAX_ENTRIES = 10000 and CACHE_EVICT_BATCH = 2000
' were declared in modConstants but eviction was never implemented.
' The Dictionary would grow without bound, potentially consuming
' hundreds of MB during long sessions.
'
' Fix: after each Put, if the dictionary has exceeded CACHE_MAX_ENTRIES,
' remove the oldest CACHE_EVICT_BATCH keys (FIFO via Keys array).
' Dictionary.Keys returns keys in insertion order in Scripting.Dictionary,
' making index-0 the oldest entry.
'
' FIX CACHE-02: IsArray guard on Keys() result; explicit nCount <= 0 exit
' before the eviction loop.
'
' v9.13.1: split into orchestrator (this procedure) + CACHE_EvictBatch.
' v9.18.0: ByVal dCache (Object is a reference already).
Private Sub CACHE_EvictIfNeeded(ByVal dCache As Object)
    If dCache Is Nothing Then Exit Sub

    Dim nCount As Long
    On Error Resume Next
    nCount = dCache.Count
    On Error GoTo 0
    If nCount <= 0 Then Exit Sub
    If nCount <= CACHE_MAX_ENTRIES Then Exit Sub

    CACHE_EvictBatch dCache, nCount
End Sub

' CACHE_Clear   empty both caches without recreating them.
'   Prefer RemoveAll over Set = Nothing + CreateObject: keeps CompareMode
'   = vbTextCompare and avoids a reallocation. Referenced by modState.
'   STATE_Reset if the session teardown wants to drop cached data.
Public Sub CACHE_Clear()
    On Error Resume Next
    If Not g_dPageCache Is Nothing Then g_dPageCache.RemoveAll
    On Error GoTo 0

    On Error Resume Next
    If Not g_dSheetCache Is Nothing Then g_dSheetCache.RemoveAll
    On Error GoTo 0
End Sub

' CACHE_EvictBatch - remove the oldest CACHE_EVICT_BATCH keys.
'   nCount is passed in (already validated > CACHE_MAX_ENTRIES > 0).
'   Keys() may return a non-array if the dictionary is in a bad state;
'   guarded. nEvict is clamped to nCount so a misconfigured
'   CACHE_EVICT_BATCH larger than the dictionary cannot loop past the end.
'   v9.18.0: ByVal dCache.
Private Sub CACHE_EvictBatch(ByVal dCache As Object, ByVal nCount As Long)
    Dim aKeys() As Variant
    Dim i       As Long
    Dim nEvict  As Long

    On Error Resume Next
    aKeys = dCache.Keys
    On Error GoTo 0
    If Not IsArray(aKeys) Then Exit Sub

    nEvict = CACHE_EVICT_BATCH
    If nEvict > nCount Then nEvict = nCount
    If nEvict <= 0 Then Exit Sub

    For i = 0 To nEvict - 1
        On Error Resume Next
        dCache.Remove aKeys(i)
        On Error GoTo 0
    Next i

    Dim nRemaining As Long
    nRemaining = 0
    On Error Resume Next
    nRemaining = dCache.Count
    On Error GoTo 0

    PRJ_LogInfo "modCache.CACHE_EvictBatch", "Evicted " & nEvict & " entries, remaining " & nRemaining
End Sub

' END: modCache v9.19.9

