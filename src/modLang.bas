Attribute VB_Name = "modLang"
Option Explicit
' Module    : modLang
' Purpose   : Batch PDF Publisher localization engine — flat key->string
'             table with EN fallback, host auto-detection, and registry
'             override. Provides L(), L2(), L3(), LN(), U(), and the
'             LANG_* code normalization helpers.
' Reference : Appendix A.13
' Version   : 9.19.10 (Inventor 2027)
'
' v9.19.10 (dangling key cleanup):
'   Removed 10 unused tip.* keys from all 13 languages:
'     tip.browse, tip.txtpath, tip.copy, tip.back,
'     tip.filter.all, tip.filter.err, tip.filter.skip,
'     tip.filter.ok, tip.title, tip.progress
'   PRJ_ApplyTooltips (modUIHelpers v9.19.9) no longer consumes
'   them — they were superseded when the tooltip list was
'   trimmed from 15 to 12 entries.
'   Removed msg.apprentice.missing from L_FillEN. Superseded by
'   modHelp (HELP_ApprenticeShortMsg / HELP_ApprenticeREADMEBody
'   in 13 languages). Only EN had this key; the other 12
'   languages were silently falling back to EN at runtime.
'   Result: EN and all 12 other languages have identical key
'   sets (178 keys each). Zero dangling keys.
'
' v9.19.9 (Badge redesign — 3 new keys):
'   Added scan.sum.to_create, scan.sum.to_recreate and
'   scan.sum.force_suffix to all 13 languages. Consumed by
'   modUIHelpers.SC_BuildBadgeText (v9.19.9) for the visible
'   lblSum caption. The diagnostic keys (scan.sum.missing,
'   scan.sum.stale, scan.sum.mismatch, scan.sum.all.current)
'   remain — SC_BuildBadgeTooltip and PRJ_GetAuditSummary still
'   read them.
'
' v9.19.6 (Version alignment + header standardization):
'   Header converted from Style A (free-form) to Style B
'   (Module/Purpose/Reference/Version). Serbian comments in
'   LANG_Init translated to English. Requires section corrected:
'   the module depends on REG_Get (modState) and
'   ThisApplication.LanguageCode (Inventor host). No functional
'   changes.
'
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.19.5: Header sync to PRJ_VERSION. No functional changes.
'
' SIZE NOTE: This module exceeds the 800 LOC guideline (section 1.10)
' and is documented as a justified exception.
'
' Justification:
'   1. Single responsibility: flat key->string localization table.
'   2. Splitting into 13 per-language modules would move the project
'      from 36 to 49 modules, triggering over-modularization (1.1).
'   3. Analogous to an embedded .resx resource file in .NET.
'   4. Adding a new language = adding one Public Sub L_FillXX.
'   5. Data is contiguous per language; navigation is trivial.
'
' This is a deliberate architectural choice, not debt.
'
' v9.19.3 (Localized theme names):
'   Added theme.inventor key to all 13 languages. Same value
'   ("Inventor") everywhere — brand name, not translated. Goes
'   through L() for uniformity with the other theme.* keys
'   consumed by PRJ_UpdateThemeLabel (modTheme v9.19.3).
'
' v9.19.2 (Date/time locale + missing keys + CJK encoding fix):
'   1. Removed report.date.format and report.date.timeonly from all
'      13 languages. Date/time now reads from Windows Region settings
'      via modFileInfo.PRJ_FormatDate / PRJ_FormatNow. Per-language
'      hardcoded templates ignored the user's actual Windows locale.
'   2. Added report.summary.full and report.skip.empty to all 13
'      languages. modReportFilter.RP_BuildSummaryText referenced
'      both, but neither key existed anywhere — summary line showed
'      "??report.summary.full??" in every language.
'   3. L_FillJA / L_FillKO / L_FillZH / L_FillZHT rewritten to use
'      U("...\uXXXX...") exclusively. Raw CJK literals survived the
'      .bas save only as "?" on CP1252 hosts, leaving JA/KO/ZH/ZHT
'      UIs showing "PDF ????" instead of translated captions.
'
' v9.19.1 (UI text consolidation):
'   7 keys added in all 13 languages: chk.force, scan.reset.sure,
'   scan.reset.done, report.form.title, report.hdr.sub,
'   report.hdr.force_short, report.hdr.audit.
'   4 dead keys removed: scan.sum.models, scan.sum.all.covered,
'   scan.audit.sum.with, scan.audit.sum.without.
'
' v9.19.0 (Audit summary narrative):
'   Replaced summary.hdr.assy / summary.hdr.part / summary.hdr.idw
'   with summary.hdr.models_full. Reworded with_idw/no_idw.
'
' v9.18.7 (§18.1 Externalization):
'   9 new keys: btn.back/export/retry/reset, report.filter.*.
'
' v9.18.0 (§4.11, §4.15):
'   7 new keys: errdlg.label.*, pdf.err.prepare/header/unknown.
'
' v9.18.0 (§1.17 Resource Separation):
'   6 new keys: scan.title.no_folder, summary.hdr.*
'
' v9.17.0: tip.title, tip.progress, tip.search, msg.confirm.reset.
' v9.13.9: LANG_SetLanguage + registry override via REG_KEY_LANG.
'
' Requires  : REG_Get (modState).
'             ThisApplication.LanguageCode (Inventor host).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.
'             The folder picker dialog uses L("scan.browse_title"),
'             added to all 13 L_FillXX procedures in modLang v9.19.6.

Private m_oDict As Object
Private m_oEN   As Object
Private m_sLang As String

Public Sub LANG_Init()
    Dim sCode As String
    Dim sOverride As String
    Dim oDict As Object

    ' 1) Check registry override
    On Error Resume Next
    sOverride = LCase$(REG_Get(REG_KEY_LANG, ""))
    On Error GoTo 0

    If Len(sOverride) > 0 Then
        If sOverride = "zh-cn" Then
            m_sLang = "zh"
        ElseIf sOverride = "zh-tw" Then
            m_sLang = "zht"
        Else
            m_sLang = Left$(sOverride, 3)
        End If
    Else
        ' 2) Fallback to host auto-detection
        On Error Resume Next
        sCode = LCase$(ThisApplication.LanguageCode)
        On Error GoTo 0
        If Len(sCode) = 0 Then sCode = "en"
        If sCode = "zh-cn" Then
            m_sLang = "zh"
        ElseIf sCode = "zh-tw" Then
            m_sLang = "zht"
        Else
            m_sLang = Left$(sCode, 2)
        End If
    End If

    ' Populate EN (fallback)
    Set m_oEN = CreateObject("Scripting.Dictionary")
    m_oEN.CompareMode = vbTextCompare
    L_FillEN m_oEN

    If m_sLang = "en" Then
        Set m_oDict = m_oEN
        Exit Sub
    End If

    Set oDict = CreateObject("Scripting.Dictionary")
    oDict.CompareMode = vbTextCompare
    Select Case m_sLang
        Case "de":  L_FillDE oDict
        Case "fr":  L_FillFR oDict
        Case "it":  L_FillIT oDict
        Case "es":  L_FillES oDict
        Case "pt":  L_FillPT oDict
        Case "ru":  L_FillRU oDict
        Case "pl":  L_FillPL oDict
        Case "cs":  L_FillCS oDict
        Case "ja":  L_FillJA oDict
        Case "ko":  L_FillKO oDict
        Case "zh":  L_FillZH oDict
        Case "zht": L_FillZHT oDict
        Case Else:  Set oDict = m_oEN
    End Select
    Set m_oDict = oDict
End Sub

Public Sub LANG_SetLanguage(ByVal sCode As String)
    Dim sNorm As String
    sNorm = LCase$(Trim$(sCode))

    If Len(sNorm) = 0 Then
        m_sLang = ""
        Set m_oDict = Nothing
        Set m_oEN = Nothing
        LANG_Init
        Exit Sub
    End If

    If sNorm = "zh-cn" Then sNorm = "zh"
    If sNorm = "zh-tw" Then sNorm = "zht"

    Select Case sNorm
        Case "en", "de", "fr", "it", "es", "pt"
        Case "ru", "pl", "cs", "ja", "ko", "zh", "zht"
        Case Else
            sNorm = "en"
    End Select

    m_sLang = sNorm
    Set m_oDict = Nothing
    Set m_oEN = Nothing
    LANG_Init
End Sub

Public Function L(sKey As String) As String
    If m_oDict Is Nothing Then LANG_Init
    If m_oDict Is Nothing Then
        L = "??" & sKey & "??"
        Exit Function
    End If
    If m_oDict.Exists(sKey) Then
        L = m_oDict(sKey)
        Exit Function
    End If
    If Not m_oEN Is Nothing Then
        If m_oEN.Exists(sKey) Then
            L = m_oEN(sKey)
            Exit Function
        End If
    End If
    L = "??" & sKey & "??"
End Function

Public Function L2(sKey As String, sA1 As String) As String
    L2 = Replace$(L(sKey), "{0}", sA1)
End Function

Public Function L3(sKey As String, sA1 As String, sA2 As String) As String
    Dim s As String
    s = Replace$(L(sKey), "{0}", sA1)
    L3 = Replace$(s, "{1}", sA2)
End Function

Public Function LN(sKey As String, ParamArray aArgs() As Variant) As String
    Dim s As String
    Dim i As Long
    Dim nLo As Long
    Dim nHi As Long

    s = L(sKey)
    nLo = LBound(aArgs)
    nHi = UBound(aArgs)
    If nHi < nLo Then
        LN = s
        Exit Function
    End If

    For i = nLo To nHi
        s = Replace$(s, "{" & i & "}", CStr(aArgs(i)))
    Next i
    LN = s
End Function

Public Function U(sInput As String) As String
    Dim i As Long, nLen As Long, nOut As Long
    Dim aBuf() As String, sHex As String, nCode As Long
    nLen = Len(sInput)
    If nLen = 0 Then U = "": Exit Function
    ReDim aBuf(0 To nLen - 1)
    nOut = 0: i = 1
    Do While i <= nLen
        If Mid$(sInput, i, 2) = "\u" And i + 5 <= nLen Then
            sHex = Mid$(sInput, i + 2, 4)
            nCode = 0
            On Error Resume Next
            nCode = CLng("&H" & sHex)
            On Error GoTo 0
            If nCode > 0 Then
                aBuf(nOut) = ChrW$(nCode): nOut = nOut + 1: i = i + 6
            Else
                aBuf(nOut) = Mid$(sInput, i, 1): nOut = nOut + 1: i = i + 1
            End If
        Else
            aBuf(nOut) = Mid$(sInput, i, 1): nOut = nOut + 1: i = i + 1
        End If
    Loop
    U = Join(aBuf, "")
    If nOut < nLen Then U = Left$(U, nOut)
End Function

Public Function LANG_ActiveCode() As String
    LANG_ActiveCode = m_sLang
End Function

Public Function LANG_SupportedCodes() As Variant
    LANG_SupportedCodes = Array("", "en", "de", "fr", "it", "es", "pt", _
        "ru", "pl", "cs", "ja", "ko", "zh", "zht")
End Function

Public Function LANG_NormalizeCode(ByVal sCode As String) As String
    Dim s As String
    s = LCase$(Trim$(sCode))
    If s = "zh-cn" Then
        LANG_NormalizeCode = "zh"
    ElseIf s = "zh-tw" Then
        LANG_NormalizeCode = "zht"
    ElseIf Len(s) > 2 Then
        LANG_NormalizeCode = Left$(s, 2)
    Else
        LANG_NormalizeCode = s
    End If
End Function

Public Function LANG_IndexOfCode(ByVal sCode As String) As Long
    Dim v As Variant
    Dim i As Long
    LANG_IndexOfCode = -1
    v = LANG_SupportedCodes()
    For i = LBound(v) To UBound(v)
        If LCase$(sCode) = CStr(v(i)) Then
            LANG_IndexOfCode = i
            Exit Function
        End If
    Next i
End Function

Private Sub L_FillEN(d As Object)
    d.Add "title", "Batch PDF Publisher"
    d.Add "btn.run", "Run"
    d.Add "btn.cancel", "Cancel"
    d.Add "btn.browse", "Browse..."
    d.Add "btn.copy", "Copy"
    d.Add "btn.help", "Help"
    d.Add "btn.settings", "Settings"
    d.Add "btn.back", "Back"
    d.Add "btn.export", "Export"
    d.Add "btn.retry", "Retry"
    d.Add "btn.reset", "Reset"
    d.Add "tab.scan", "Scan"
    d.Add "tab.settings", "Settings"
    d.Add "tab.log", "Log"
    d.Add "settings.lang", "Language:"
    d.Add "settings.lang.auto", "Auto (host)"
    d.Add "settings.loglevel", "Log level:"
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", "Warning"
    d.Add "settings.loglevel.error", "Error only"
    d.Add "lbl.folder", "Source folder:"
    d.Add "lbl.subfolders", "Include subfolders"
    d.Add "chk.force", "Force recreate"
    d.Add "chk.audit", "Audit drawings"
    d.Add "lbl.options", "Output options:"
    d.Add "opt.single", "Single PDF (combined)"
    d.Add "opt.multi", "Individual files"
    d.Add "chk.open", "Open folder when done"
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", "Light"
    d.Add "theme.dark", "Dark"
    d.Add "theme.system", "System"
    d.Add "msg.select.folder", "Select a folder to scan."
    d.Add "msg.no.idw", "No Inventor drawings (.idw) found in the selected path."
    d.Add "msg.confirm.start", "Start processing {0} files?"
    d.Add "msg.process.done", "Process completed successfully."
    d.Add "msg.error.title", "System Error"
    d.Add "log.start", "Session started: {0}"
    d.Add "log.end", "Session ended: {0}"
    d.Add "log.save", "Log saved to: {0}"
    d.Add "scan.sum.drawings", "{0} drawings"
    d.Add "scan.sum.subfolders", "including subfolders"
    d.Add "scan.sum.ok", "{0} current"
    d.Add "scan.sum.all.current", "all current"
    d.Add "scan.sum.no.idw", "No drawings found"
    d.Add "scan.sum.missing", "{0} missing"
    d.Add "scan.sum.stale", "{0} stale"
    d.Add "scan.sum.mismatch", "{0} mismatched"
    d.Add "scan.sum.to_create", "{0} to create"
    d.Add "scan.sum.to_recreate", "{0} to recreate"
    d.Add "scan.sum.force_suffix", "(force)"
    d.Add "scan.validate.error", "Error validating folder."
    d.Add "scan.sum.errors", "{0} errors"
    d.Add "scan.no.folder", "No folder selected."
    d.Add "scan.no.folder.exists", "Folder does not exist."
    d.Add "scan.path.ok", "Folder found - scanning..."
    d.Add "scan.path.too.long", "Path exceeds {0} character limit."
    d.Add "progress.processing", "Processing..."
    d.Add "progress.init", "Initializing..."
    d.Add "progress.cancel", "Cancel"
    d.Add "progress.cancelling", "Cancelling..."
    d.Add "scan.models.detail", "With drawing: {0} | Without: {1}"
    d.Add "status.idle", "Ready"
    d.Add "status.scanning", "Scanning folder..."
    d.Add "status.publishing", "Publishing: {0} of {1}..."
    d.Add "pdf.err.open", "Could not open IDW file."
    d.Add "pdf.err.pagematch", "PDF has {0} pages, expected {1}."
    d.Add "pdf.err.noext", "Invalid file path (no extension): {0}"
    d.Add "pdf.skipped.exists", "Skipped (exists)"
    d.Add "pdf.skipped.exists.pages", "Skipped (exists) - {0}/{1} pages"
    d.Add "pdf.err.addin", "PDF Translator AddIn not found."
    d.Add "pdf.err.nofile", "PDF export produced no file."
    d.Add "pdf.err.export", "Export failed: {0}"
    d.Add "pdf.err.bak.delete", "Cannot remove stale backup: {0}"
    d.Add "pdf.err.bak.move", "Cannot back up existing PDF: {0}"
    d.Add "pdf.err.write", "Cannot write PDF: {0}"
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 Export Report")
    d.Add "report.hdr.mode", "Mode: {0}"
    d.Add "report.hdr.folder", "Folder: {0}"
    d.Add "report.hdr.subfolders", "Subfolders"
    d.Add "report.hdr.force", "Force overwrite"
    d.Add "report.hdr.generated", "Generated: {0}"
    d.Add "report.state.on", "On"
    d.Add "report.state.off", "Off"
    d.Add "report.hdr.pdf.section", "PDF EXPORT"
    d.Add "report.hdr.pdf.detail", "Details:"
    d.Add "report.created", "Created"
    d.Add "report.overwritten", "Overwritten"
    d.Add "report.skipped", "Skipped"
    d.Add "report.errors", "Errors"
    d.Add "report.time", "Time"
    d.Add "report.empty", "empty"
    d.Add "report.pages", "pages"
    d.Add "report.no.views", "No views"
    d.Add "report.tag.created", "[Created]"
    d.Add "report.tag.overwritten", "[Overwritten]"
    d.Add "report.tag.skipped", "[Skipped]"
    d.Add "report.tag.error", "[Error]"
    d.Add "report.tag.hasdrawing", "[Has Drawing]"
    d.Add "report.tag.nodrawing", "[No Drawing]"
    d.Add "report.hdr.audit.section", "DRAWING AUDIT"
    d.Add "report.hdr.audit.detail", "Details:"
    d.Add "report.audit.title", "Drawing Audit"
    d.Add "errdlg.retry.prompt", "Retry = try again | Ignore = skip | Abort = stop"
    d.Add "tip.run", "Start batch PDF export"
    d.Add "tip.theme", "Cycle through themes (Inventor / System / Light / Dark)"
    d.Add "tip.reset", "Reset all options to defaults"
    d.Add "tip.cancel", "Cancel current operation (ESC)"
    d.Add "tip.filter.nodrawing", "Show only models without a drawing"
    d.Add "tip.list", "Double-click a row to open the PDF; click to see details"
    d.Add "tip.subfolders", "Recursively scan subfolders (skips OldVersions)"
    d.Add "tip.force", "Recreate PDF even if it is newer than the IDW"
    d.Add "tip.audit", "Audit models (.ipt/.iam) for missing sibling drawings"
    d.Add "tip.retry", "Retry all files that ended in error"
    d.Add "tip.export", "Save report to file and open in Notepad"
    d.Add "report.msg.close_blocked", "Cannot close during processing."
    d.Add "report.msg.open_failed", "Cannot open: file not found or action blocked."
    d.Add "report.msg.select_row", "Select a row for details  " & Chr(183) & "  Double-click to open PDF"
    d.Add "report.msg.retry.none", "Retry: no files in error state."
    d.Add "report.msg.retry.error", "Retry error: {0}"
    d.Add "report.msg.retry.partial", "Retry: succeeded {0}, still in error {1}."
    d.Add "report.msg.retry.all_fixed", "Retry: succeeded {0}, all fixed."
    d.Add "report.msg.copied", "Copied!"
    d.Add "report.msg.exported", "Report opened in Notepad."
    d.Add "report.msg.export_failed", "Failed to save report."
    d.Add "reset.msg.all_defaults", "All settings reset to defaults."
    d.Add "reset.lbl.registry", "Registry"
    d.Add "reset.lbl.sandbox", "Sandbox"
    d.Add "reset.lbl.state", "State"
    d.Add "reset.state.deleted", "Deleted"
    d.Add "reset.state.reinitialized", "Reinitialized"
    d.Add "reset.msg.failed", "Reset failed"
    d.Add "reset.msg.registry_complete", "Registry reset complete."
    d.Add "reset.msg.sandbox_complete", "Sandbox reset complete"
    d.Add "scan.badge.tooltip", "Total: {0} | Current: {1} | Missing: {2} | Stale: {3} | Mismatched: {4} | Errors: {5}"
    d.Add "report.msg.opening", "Opening IDW..."
    d.Add "report.msg.updating", "Updating views..."
    d.Add "report.msg.exporting", "Exporting PDF..."
    d.Add "msg.no.write.access", "Cannot write to sandbox folder. Check permissions:"
    d.Add "msg.pdf.temp_missing", "Temp PDF disappeared during export (possible antivirus quarantine or file lock): {0}"
    d.Add "msg.batch.abort.network", "Batch aborted: {0} consecutive network failures. Check connection and retry."
    d.Add "msg.already.running", "Batch PDF Publisher is already running. Close the other session and try again."
    d.Add "report.search.placeholder", "Search path..."
    d.Add "report.search.label", "Search:"
    d.Add "scan.collecting", "Collecting files..."
    d.Add "scan.analysing", "Analysing: {0} / {1}..."
    d.Add "scan.auditing.start", "Auditing models..."
    d.Add "scan.auditing", "Auditing: {0} / {1}..."
    d.Add "scan.cancelled", "Cancelled"
    d.Add "tip.search", "Filter report rows by path substring"
    d.Add "msg.confirm.reset", "Reset all options to defaults?"
    d.Add "scan.disk.insufficient", "Insufficient disk space. Free: {0}"
    d.Add "scan.title.no_folder", "Please select folder"
    d.Add "summary.hdr.models_full", "{0} models ({1} assemblies, {2} parts)"
    d.Add "summary.hdr.with_idw", "{0} covered"
    d.Add "summary.hdr.no_idw", "{0} missing"
    d.Add "errdlg.label.context", "Context"
    d.Add "errdlg.label.time", "Time"
    d.Add "errdlg.label.log", "Log"
    d.Add "errdlg.label.stack", "Stack"
    d.Add "pdf.err.prepare", "Prepare failed: {0}"
    d.Add "pdf.err.header", "Exported temp PDF failed header validation"
    d.Add "pdf.err.unknown", "Unknown error"
    d.Add "report.filter.all", "All"
    d.Add "report.filter.err", "Errors"
    d.Add "report.filter.skip", "Skipped"
    d.Add "report.filter.ok", "OK"
    d.Add "report.filter.noidw", "No .idw"
    d.Add "scan.reset.sure", "Sure?"
    d.Add "scan.reset.done", "Done"
    d.Add "report.form.title", "Processing Report"
    d.Add "report.hdr.sub", "Subfolders"
    d.Add "report.hdr.force_short", "Force"
    d.Add "report.hdr.audit", "Audit"
    d.Add "report.summary.full", U("{0} created \u00b7 {1} overwritten \u00b7 {2} skipped \u00b7 {3} errors \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} empty)")
    d.Add "scan.browse_title", "Select source folder"
End Sub

Private Sub L_FillDE(d As Object)
    d.Add "title", "Stapel-PDF-Publisher"
    d.Add "btn.run", "STARTEN"
    d.Add "btn.cancel", "Abbrechen"
    d.Add "btn.browse", "Durchsuchen..."
    d.Add "btn.copy", "Kopieren"
    d.Add "btn.help", "Hilfe"
    d.Add "btn.settings", "Einstellungen"
    d.Add "btn.back", U("Zur\u00fcck")
    d.Add "btn.export", "Export"
    d.Add "btn.retry", "Wiederholen"
    d.Add "btn.reset", U("Zur\u00fccksetzen")
    d.Add "tab.scan", "Scannen"
    d.Add "tab.settings", "Einstellungen"
    d.Add "tab.log", "Protokoll"
    d.Add "settings.lang", "Sprache:"
    d.Add "settings.lang.auto", "Automatisch (Host)"
    d.Add "settings.loglevel", "Protokollstufe:"
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", "Warnung"
    d.Add "settings.loglevel.error", "Nur Fehler"
    d.Add "lbl.folder", "Quellordner:"
    d.Add "lbl.subfolders", "Unterordner einbeziehen"
    d.Add "chk.force", "Neuerstellen erzwingen"
    d.Add "chk.audit", U("Zeichnungen pr\u00fcfen")
    d.Add "lbl.options", "Ausgabeoptionen:"
    d.Add "opt.single", "Einzelne PDF (kombiniert)"
    d.Add "opt.multi", "Einzelne Dateien"
    d.Add "chk.open", U("Ordner nach Abschluss \u00f6ffnen")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", "Hell"
    d.Add "theme.dark", "Dunkel"
    d.Add "theme.system", "System"
    d.Add "msg.select.folder", U("Bitte einen Ordner zum Scannen ausw\u00e4hlen.")
    d.Add "msg.no.idw", U("Keine Inventor-Zeichnungen (.idw) im gew\u00e4hlten Pfad gefunden.")
    d.Add "msg.confirm.start", "Verarbeitung von {0} Dateien starten?"
    d.Add "msg.process.done", "Vorgang erfolgreich abgeschlossen."
    d.Add "msg.error.title", "Systemfehler"
    d.Add "log.start", "Sitzung gestartet: {0}"
    d.Add "log.end", "Sitzung beendet: {0}"
    d.Add "log.save", "Protokoll gespeichert unter: {0}"
    d.Add "scan.sum.drawings", "{0} Zeichnungen"
    d.Add "scan.sum.subfolders", "inkl. Unterordner"
    d.Add "scan.sum.ok", "{0} aktuell"
    d.Add "scan.sum.all.current", "alle aktuell"
    d.Add "scan.sum.no.idw", "Keine Zeichnungen gefunden"
    d.Add "scan.sum.missing", "{0} fehlend"
    d.Add "scan.sum.stale", "{0} veraltet"
    d.Add "scan.sum.mismatch", "{0} abweichend"
    d.Add "scan.sum.to_create", "{0} neu erstellen"
    d.Add "scan.sum.to_recreate", "{0} neu erzeugen"
    d.Add "scan.sum.force_suffix", "(erzwungen)"
    d.Add "scan.validate.error", "Fehler bei der Ordnervalidierung."
    d.Add "scan.sum.errors", "{0} Fehler"
    d.Add "scan.no.folder", U("Kein Ordner ausgew\u00e4hlt.")
    d.Add "scan.no.folder.exists", "Ordner existiert nicht."
    d.Add "scan.path.ok", U("Ordner gefunden - wird gescannt...")
    d.Add "scan.path.too.long", U("Pfad \u00fcberschreitet das Limit von {0} Zeichen.")
    d.Add "progress.processing", "Verarbeitung..."
    d.Add "progress.init", "Initialisierung..."
    d.Add "progress.cancel", "Abbrechen"
    d.Add "progress.cancelling", "Wird abgebrochen..."
    d.Add "scan.models.detail", "Mit Zeichnung: {0} | Ohne: {1}"
    d.Add "status.idle", "Bereit"
    d.Add "status.scanning", "Ordner wird gescannt..."
    d.Add "status.publishing", "Publizieren: {0} von {1}..."
    d.Add "pdf.err.open", U("IDW-Datei konnte nicht ge\u00f6ffnet werden.")
    d.Add "pdf.err.pagematch", "PDF hat {0} Seiten, erwartet {1}."
    d.Add "pdf.err.noext", U("Ung\u00fcltiger Dateipfad (keine Erweiterung): {0}")
    d.Add "pdf.skipped.exists", U("\u00dcbersprungen (vorhanden)")
    d.Add "pdf.skipped.exists.pages", U("\u00dcbersprungen (vorhanden) - {0}/{1} Seiten")
    d.Add "pdf.err.addin", "PDF-Translator-AddIn nicht gefunden."
    d.Add "pdf.err.nofile", "PDF-Export hat keine Datei erzeugt."
    d.Add "pdf.err.export", "Export fehlgeschlagen: {0}"
    d.Add "pdf.err.bak.delete", U("Altes Backup kann nicht gel\u00f6scht werden: {0}")
    d.Add "pdf.err.bak.move", "Vorhandene PDF kann nicht gesichert werden: {0}"
    d.Add "pdf.err.write", "PDF kann nicht geschrieben werden: {0}"
    d.Add "report.hdr.title", U("Batch-PDF-Publisher v{0} \u2014 Exportbericht")
    d.Add "report.hdr.mode", "Modus: {0}"
    d.Add "report.hdr.folder", "Ordner: {0}"
    d.Add "report.hdr.subfolders", "Unterordner"
    d.Add "report.hdr.force", U("\u00dcberschreiben erzwingen")
    d.Add "report.hdr.generated", "Erstellt: {0}"
    d.Add "report.state.on", "Ein"
    d.Add "report.state.off", "Aus"
    d.Add "report.hdr.pdf.section", "PDF-EXPORT"
    d.Add "report.hdr.pdf.detail", "Details:"
    d.Add "report.created", "Erstellt"
    d.Add "report.overwritten", U("\u00dcberschrieben")
    d.Add "report.skipped", U("\u00dcbersprungen")
    d.Add "report.errors", "Fehler"
    d.Add "report.time", "Zeit"
    d.Add "report.empty", "leer"
    d.Add "report.pages", "Seiten"
    d.Add "report.no.views", "Keine Ansichten"
    d.Add "report.tag.created", "[Erstellt]"
    d.Add "report.tag.overwritten", U("[\u00dcberschrieben]")
    d.Add "report.tag.skipped", U("[\u00dcbersprungen]")
    d.Add "report.tag.error", "[Fehler]"
    d.Add "report.tag.hasdrawing", "[Hat Zeichnung]"
    d.Add "report.tag.nodrawing", "[Keine Zeichnung]"
    d.Add "report.hdr.audit.section", "ZEICHNUNGS-AUDIT"
    d.Add "report.hdr.audit.detail", "Details:"
    d.Add "report.audit.title", "Zeichnungs-Audit"
    d.Add "errdlg.retry.prompt", U("Wiederholen = erneut versuchen | Ignorieren = \u00fcberspringen | Abbrechen = stoppen")
    d.Add "tip.run", "Batch-PDF-Export starten"
    d.Add "tip.theme", U("Themes durchschalten (Inventor / System / Hell / Dunkel)")
    d.Add "tip.reset", U("Alle Optionen auf Standard zur\u00fccksetzen")
    d.Add "tip.cancel", U("Aktuellen Vorgang abbrechen (ESC)")
    d.Add "tip.filter.nodrawing", U("Nur Modelle ohne Zeichnung anzeigen")
    d.Add "tip.list", U("Doppelklick \u00f6ffnet PDF; Klick zeigt Details")
    d.Add "tip.subfolders", U("Unterordner rekursiv scannen (OldVersions wird \u00fcbersprungen)")
    d.Add "tip.force", U("PDF neu erstellen, auch wenn neuer als IDW")
    d.Add "tip.audit", U("Modelle (.ipt/.iam) auf fehlende Zeichnungen pr\u00fcfen")
    d.Add "tip.retry", U("Alle fehlerhaften Dateien erneut versuchen")
    d.Add "tip.export", U("Bericht speichern und in Notepad \u00f6ffnen")
    d.Add "report.msg.close_blocked", U("Schlie\u00dfen w\u00e4hrend der Verarbeitung nicht m\u00f6glich.")
    d.Add "report.msg.open_failed", U("\u00d6ffnen nicht m\u00f6glich: Datei nicht gefunden oder Aktion blockiert.")
    d.Add "report.msg.select_row", U("Zeile ausw\u00e4hlen f\u00fcr Details  \u00b7  Doppelklick zum \u00d6ffnen der PDF")
    d.Add "report.msg.retry.none", "Wiederholen: keine Dateien im Fehlerzustand."
    d.Add "report.msg.retry.error", "Wiederholungsfehler: {0}"
    d.Add "report.msg.retry.partial", "Wiederholen: {0} erfolgreich, {1} noch im Fehler."
    d.Add "report.msg.retry.all_fixed", "Wiederholen: {0} erfolgreich, alle behoben."
    d.Add "report.msg.copied", "Kopiert!"
    d.Add "report.msg.exported", U("Bericht in Notepad ge\u00f6ffnet.")
    d.Add "report.msg.export_failed", "Bericht konnte nicht gespeichert werden."
    d.Add "reset.msg.all_defaults", "Alle Einstellungen auf Standard zurueckgesetzt."
    d.Add "reset.lbl.registry", "Registrierung"
    d.Add "reset.lbl.sandbox", "Sandbox"
    d.Add "reset.lbl.state", "Status"
    d.Add "reset.state.deleted", "Geloescht"
    d.Add "reset.state.reinitialized", "Neu initialisiert"
    d.Add "reset.msg.failed", "Zuruecksetzen fehlgeschlagen"
    d.Add "reset.msg.registry_complete", "Registrierung zurueckgesetzt."
    d.Add "reset.msg.sandbox_complete", "Sandbox zurueckgesetzt"
    d.Add "scan.badge.tooltip", "Gesamt: {0} | Aktuell: {1} | Fehlend: {2} | Veraltet: {3} | Abweichend: {4} | Fehler: {5}"
    d.Add "report.msg.opening", U("IDW wird ge\u00f6ffnet...")
    d.Add "report.msg.updating", U("Ansichten werden aktualisiert...")
    d.Add "report.msg.exporting", U("PDF wird exportiert...")
    d.Add "msg.no.write.access", U("Kein Schreibzugriff auf Sandbox-Ordner. Pr\u00fcfen Sie die Berechtigungen:")
    d.Add "msg.pdf.temp_missing", U("Tempor\u00e4re PDF ist w\u00e4hrend des Exports verschwunden (m\u00f6glicherweise Antivirus-Quarant\u00e4ne oder Dateisperre): {0}")
    d.Add "msg.batch.abort.network", U("Batch abgebrochen: {0} aufeinanderfolgende Netzwerkfehler. Pr\u00fcfen Sie die Verbindung und versuchen Sie es erneut.")
    d.Add "msg.already.running", U("Batch PDF Publisher l\u00e4uft bereits. Schlie\u00dfen Sie die andere Sitzung und versuchen Sie es erneut.")
    d.Add "report.search.placeholder", "Pfad suchen..."
    d.Add "report.search.label", "Suche:"
    d.Add "scan.collecting", "Dateien werden gesammelt..."
    d.Add "scan.analysing", "Analyse: {0} / {1}..."
    d.Add "scan.auditing.start", "Modelle werden geprueft..."
    d.Add "scan.auditing", "Audit: {0} / {1}..."
    d.Add "scan.cancelled", "Abgebrochen"
    d.Add "tip.search", "Berichtszeilen nach Pfad filtern"
    d.Add "msg.confirm.reset", U("Alle Optionen auf Standard zur\u00fccksetzen?")
    d.Add "scan.disk.insufficient", U("Nicht gen\u00fcgend Speicherplatz. Frei: {0}")
    d.Add "scan.title.no_folder", "Bitte Ordner ausw\u00e4hlen"
    d.Add "summary.hdr.models_full", "{0} Modelle ({1} Baugruppen, {2} Bauteile)"
    d.Add "summary.hdr.with_idw", "{0} mit Zeichnung"
    d.Add "summary.hdr.no_idw", "{0} ohne Zeichnungen"
    d.Add "errdlg.label.context", "Kontext"
    d.Add "errdlg.label.time", "Zeit"
    d.Add "errdlg.label.log", "Protokoll"
    d.Add "errdlg.label.stack", "Stapel"
    d.Add "pdf.err.prepare", U("Vorbereitung fehlgeschlagen: {0}")
    d.Add "pdf.err.header", U("Exportierte tempor\u00e4re PDF hat Header-Validierung nicht bestanden")
    d.Add "pdf.err.unknown", "Unbekannter Fehler"
    d.Add "report.filter.all", "Alle"
    d.Add "report.filter.err", "Fehler"
    d.Add "report.filter.skip", U("\u00dcbersprungen")
    d.Add "report.filter.ok", "OK"
    d.Add "report.filter.noidw", "Ohne .idw"
    d.Add "scan.reset.sure", "Sicher?"
    d.Add "scan.reset.done", "Fertig"
    d.Add "report.form.title", "Verarbeitungsbericht"
    d.Add "report.hdr.sub", "Unterordner"
    d.Add "report.hdr.force_short", "Erzwingen"
    d.Add "report.hdr.audit", U("Pr\u00fcfen")
    d.Add "report.summary.full", U("{0} erstellt \u00b7 {1} \u00fcberschrieben \u00b7 {2} \u00fcbersprungen \u00b7 {3} Fehler \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} leer)")
    d.Add "scan.browse_title", "Quellordner auswählen"
End Sub

Private Sub L_FillFR(d As Object)
    d.Add "title", "Publieur PDF par lots"
    d.Add "btn.run", U("EX\u00c9CUTER")
    d.Add "btn.cancel", "Annuler"
    d.Add "btn.browse", "Parcourir..."
    d.Add "btn.copy", "Copier"
    d.Add "btn.help", "Aide"
    d.Add "btn.settings", U("Param\u00e8tres")
    d.Add "btn.back", "Retour"
    d.Add "btn.export", "Exporter"
    d.Add "btn.retry", U("R\u00e9essayer")
    d.Add "btn.reset", U("R\u00e9initialiser")
    d.Add "tab.scan", "Analyser"
    d.Add "tab.settings", U("Param\u00e8tres")
    d.Add "tab.log", "Journal"
    d.Add "settings.lang", "Langue :"
    d.Add "settings.lang.auto", U("Auto (h\u00f4te)")
    d.Add "settings.loglevel", "Niveau de journal :"
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", "Avertissement"
    d.Add "settings.loglevel.error", "Erreurs uniquement"
    d.Add "lbl.folder", U("Dossier source :")
    d.Add "lbl.subfolders", "Inclure les sous-dossiers"
    d.Add "chk.force", U("Forcer la recr\u00e9ation")
    d.Add "chk.audit", U("V\u00e9rifier les dessins")
    d.Add "lbl.options", U("Options de sortie :")
    d.Add "opt.single", U("PDF unique (combin\u00e9)")
    d.Add "opt.multi", "Fichiers individuels"
    d.Add "chk.open", U("Ouvrir le dossier \u00e0 la fin")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", "Clair"
    d.Add "theme.dark", "Sombre"
    d.Add "theme.system", U("Syst\u00e8me")
    d.Add "msg.select.folder", U("S\u00e9lectionnez un dossier \u00e0 analyser.")
    d.Add "msg.no.idw", U("Aucun dessin Inventor (.idw) trouv\u00e9 dans le chemin s\u00e9lectionn\u00e9.")
    d.Add "msg.confirm.start", U("D\u00e9marrer le traitement de {0} fichiers ?")
    d.Add "msg.process.done", U("Processus termin\u00e9 avec succ\u00e8s.")
    d.Add "msg.error.title", U("Erreur syst\u00e8me")
    d.Add "log.start", U("Session d\u00e9marr\u00e9e : {0}")
    d.Add "log.end", U("Session termin\u00e9e : {0}")
    d.Add "log.save", U("Journal enregistr\u00e9 dans : {0}")
    d.Add "scan.sum.drawings", "{0} dessins"
    d.Add "scan.sum.subfolders", "sous-dossiers inclus"
    d.Add "scan.sum.ok", U("{0} \u00e0 jour")
    d.Add "scan.sum.all.current", U("tous \u00e0 jour")
    d.Add "scan.sum.no.idw", U("Aucun dessin trouv\u00e9")
    d.Add "scan.sum.missing", "{0} manquants"
    d.Add "scan.sum.stale", U("{0} obsol\u00e8tes")
    d.Add "scan.sum.mismatch", "{0} divergents"
    d.Add "scan.sum.to_create", U("{0} \u00e0 cr\u00e9er")
    d.Add "scan.sum.to_recreate", U("{0} \u00e0 recr\u00e9er")
    d.Add "scan.sum.force_suffix", U("(forc\u00e9)")
    d.Add "scan.validate.error", "Erreur lors de la validation du dossier."
    d.Add "scan.sum.errors", "{0} erreurs"
    d.Add "scan.no.folder", U("Aucun dossier s\u00e9lectionn\u00e9.")
    d.Add "scan.no.folder.exists", "Le dossier n'existe pas."
    d.Add "scan.path.ok", U("Dossier trouv\u00e9 - analyse en cours...")
    d.Add "scan.path.too.long", U("Le chemin d\u00e9passe la limite de {0} caract\u00e8res.")
    d.Add "progress.processing", "Traitement en cours..."
    d.Add "progress.init", "Initialisation..."
    d.Add "progress.cancel", "Annuler"
    d.Add "progress.cancelling", "Annulation..."
    d.Add "scan.models.detail", "Avec dessin : {0} | Sans : {1}"
    d.Add "status.idle", U("Pr\u00eat")
    d.Add "status.scanning", "Analyse du dossier..."
    d.Add "status.publishing", U("Publication : {0} sur {1}...")
    d.Add "pdf.err.open", "Impossible d'ouvrir le fichier IDW."
    d.Add "pdf.err.pagematch", "Le PDF contient {0} pages, {1} attendues."
    d.Add "pdf.err.noext", "Chemin de fichier invalide (sans extension): {0}"
    d.Add "pdf.skipped.exists", U("Ignor\u00e9 (existe)")
    d.Add "pdf.skipped.exists.pages", U("Ignor\u00e9 (existe) - {0}/{1} pages")
    d.Add "pdf.err.addin", U("Compl\u00e9ment PDF Translator introuvable.")
    d.Add "pdf.err.nofile", U("L'export PDF n'a produit aucun fichier.")
    d.Add "pdf.err.export", U("\u00c9chec de l'export : {0}")
    d.Add "pdf.err.bak.delete", U("Impossible de supprimer l'ancienne sauvegarde : {0}")
    d.Add "pdf.err.bak.move", U("Impossible de sauvegarder le PDF existant : {0}")
    d.Add "pdf.err.write", U("Impossible d'\u00e9crire le PDF : {0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 Rapport d'export")
    d.Add "report.hdr.mode", U("Mode : {0}")
    d.Add "report.hdr.folder", U("Dossier : {0}")
    d.Add "report.hdr.subfolders", "Sous-dossiers"
    d.Add "report.hdr.force", U("\u00c9crasement forc\u00e9")
    d.Add "report.hdr.generated", U("G\u00e9n\u00e9r\u00e9 : {0}")
    d.Add "report.state.on", "Oui"
    d.Add "report.state.off", "Non"
    d.Add "report.hdr.pdf.section", "EXPORT PDF"
    d.Add "report.hdr.pdf.detail", U("D\u00e9tails :")
    d.Add "report.created", U("Cr\u00e9\u00e9")
    d.Add "report.overwritten", U("\u00c9cras\u00e9")
    d.Add "report.skipped", U("Ignor\u00e9")
    d.Add "report.errors", "Erreurs"
    d.Add "report.time", "Temps"
    d.Add "report.empty", "vide"
    d.Add "report.pages", "pages"
    d.Add "report.no.views", "Aucune vue"
    d.Add "report.tag.created", U("[Cr\u00e9\u00e9]")
    d.Add "report.tag.overwritten", U("[\u00c9cras\u00e9]")
    d.Add "report.tag.skipped", U("[Ignor\u00e9]")
    d.Add "report.tag.error", "[Erreur]"
    d.Add "report.tag.hasdrawing", "[A un dessin]"
    d.Add "report.tag.nodrawing", "[Pas de dessin]"
    d.Add "report.hdr.audit.section", "AUDIT DES DESSINS"
    d.Add "report.hdr.audit.detail", U("D\u00e9tails :")
    d.Add "report.audit.title", "Audit des dessins"
    d.Add "errdlg.retry.prompt", U("R\u00e9essayer = r\u00e9essayer | Ignorer = passer | Abandonner = arr\u00eater")
    d.Add "tip.run", U("D\u00e9marrer l'export PDF par lots")
    d.Add "tip.theme", U("Parcourir les th\u00e8mes (Inventor / Syst\u00e8me / Clair / Sombre)")
    d.Add "tip.reset", U("R\u00e9initialiser toutes les options")
    d.Add "tip.cancel", U("Annuler l'op\u00e9ration en cours (ESC)")
    d.Add "tip.filter.nodrawing", U("Afficher uniquement les mod\u00e8les sans dessin")
    d.Add "tip.list", U("Double-cliquez pour ouvrir le PDF ; cliquez pour les d\u00e9tails")
    d.Add "tip.subfolders", U("Analyser r\u00e9cursivement les sous-dossiers (sauf OldVersions)")
    d.Add "tip.force", U("Recr\u00e9er le PDF m\u00eame s'il est plus r\u00e9cent que l'IDW")
    d.Add "tip.audit", U("V\u00e9rifier les mod\u00e8les (.ipt/.iam) sans dessin associ\u00e9")
    d.Add "tip.retry", U("R\u00e9essayer tous les fichiers en erreur")
    d.Add "tip.export", U("Enregistrer le rapport et l'ouvrir dans Notepad")
    d.Add "report.msg.close_blocked", U("Fermeture impossible pendant le traitement.")
    d.Add "report.msg.open_failed", U("Ouverture impossible : fichier introuvable ou action bloqu\u00e9e.")
    d.Add "report.msg.select_row", U("S\u00e9lectionnez une ligne pour les d\u00e9tails  \u00b7  Double-clic pour ouvrir le PDF")
    d.Add "report.msg.retry.none", U("R\u00e9essayer : aucun fichier en erreur.")
    d.Add "report.msg.retry.error", U("Erreur de r\u00e9essai : {0}")
    d.Add "report.msg.retry.partial", U("R\u00e9essai : {0} r\u00e9ussis, {1} toujours en erreur.")
    d.Add "report.msg.retry.all_fixed", U("R\u00e9essai : {0} r\u00e9ussis, tous corrig\u00e9s.")
    d.Add "report.msg.copied", U("Copi\u00e9 !")
    d.Add "report.msg.exported", "Rapport ouvert dans Notepad."
    d.Add "report.msg.export_failed", U("\u00c9chec de l'enregistrement du rapport.")
    d.Add "reset.msg.all_defaults", "Tous les parametres reinitialises."
    d.Add "reset.lbl.registry", "Registre"
    d.Add "reset.lbl.sandbox", "Bac a sable"
    d.Add "reset.lbl.state", "Etat"
    d.Add "reset.state.deleted", "Supprime"
    d.Add "reset.state.reinitialized", "Reinitialise"
    d.Add "reset.msg.failed", "Echec de la reinitialisation"
    d.Add "reset.msg.registry_complete", "Registre reinitialise."
    d.Add "reset.msg.sandbox_complete", "Bac a sable reinitialise"
    d.Add "scan.badge.tooltip", U("Total : {0} | Actuel : {1} | Manquant : {2} | Obsol\u00e8te : {3} | Divergent : {4} | Erreurs : {5}")
    d.Add "report.msg.opening", U("Ouverture du fichier IDW...")
    d.Add "report.msg.updating", U("Mise \u00e0 jour des vues...")
    d.Add "report.msg.exporting", U("Export PDF en cours...")
    d.Add "msg.no.write.access", U("Impossible d'\u00e9crire dans le dossier sandbox. V\u00e9rifiez les autorisations :")
    d.Add "msg.pdf.temp_missing", U("Le PDF temporaire a disparu pendant l'export (quarantaine antivirus ou verrouillage de fichier possible) : {0}")
    d.Add "msg.batch.abort.network", U("Lot interrompu : {0} erreurs r\u00e9seau cons\u00e9cutives. V\u00e9rifiez la connexion et r\u00e9essayez.")
    d.Add "msg.already.running", U("Batch PDF Publisher est d\u00e9j\u00e0 en cours d'ex\u00e9cution. Fermez l'autre session et r\u00e9essayez.")
    d.Add "report.search.placeholder", U("Rechercher un chemin...")
    d.Add "report.search.label", U("Rechercher :")
    d.Add "scan.collecting", U("Collecte des fichiers...")
    d.Add "scan.analysing", "Analyse : {0} / {1}..."
    d.Add "scan.auditing.start", U("Audit des mod\u00e8les...")
    d.Add "scan.auditing", "Audit : {0} / {1}..."
    d.Add "scan.cancelled", U("Annul\u00e9")
    d.Add "tip.search", "Filtrer les lignes du rapport par chemin"
    d.Add "msg.confirm.reset", U("R\u00e9initialiser toutes les options ?")
    d.Add "scan.disk.insufficient", U("Espace disque insuffisant. Libre : {0}")
    d.Add "scan.title.no_folder", U("Veuillez s\u00e9lectionner un dossier")
    d.Add "summary.hdr.models_full", U("{0} mod\u00e8les ({1} assemblages, {2} pi\u00e8ces)")
    d.Add "summary.hdr.with_idw", "{0} avec dessin"
    d.Add "summary.hdr.no_idw", "{0} sans dessin"
    d.Add "errdlg.label.context", "Contexte"
    d.Add "errdlg.label.time", "Heure"
    d.Add "errdlg.label.log", "Journal"
    d.Add "errdlg.label.stack", "Pile"
    d.Add "pdf.err.prepare", U("Pr\u00e9paration \u00e9chou\u00e9e : {0}")
    d.Add "pdf.err.header", U("Le PDF temporaire export\u00e9 a \u00e9chou\u00e9 \u00e0 la validation d'en-t\u00eate")
    d.Add "pdf.err.unknown", "Erreur inconnue"
    d.Add "report.filter.all", "Tous"
    d.Add "report.filter.err", "Erreurs"
    d.Add "report.filter.skip", U("Ignor\u00e9s")
    d.Add "report.filter.ok", "OK"
    d.Add "report.filter.noidw", "Sans .idw"
    d.Add "scan.reset.sure", U("Confirmer ?")
    d.Add "scan.reset.done", U("Termin\u00e9")
    d.Add "report.form.title", "Rapport de traitement"
    d.Add "report.hdr.sub", "Sous-dossiers"
    d.Add "report.hdr.force_short", "Forcer"
    d.Add "report.hdr.audit", U("V\u00e9rifier")
    d.Add "report.summary.full", U("{0} cr\u00e9\u00e9s \u00b7 {1} \u00e9cras\u00e9s \u00b7 {2} ignor\u00e9s \u00b7 {3} erreurs \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} vides)")
    d.Add "scan.browse_title", U("S\u00e9lectionner le dossier source")
End Sub

Private Sub L_FillIT(d As Object)
    d.Add "title", "Pubblicatore PDF in batch"
    d.Add "btn.run", "ESEGUI"
    d.Add "btn.cancel", "Annulla"
    d.Add "btn.browse", "Sfoglia..."
    d.Add "btn.copy", "Copia"
    d.Add "btn.help", "Aiuto"
    d.Add "btn.settings", "Impostazioni"
    d.Add "btn.back", "Indietro"
    d.Add "btn.export", "Esporta"
    d.Add "btn.retry", "Riprova"
    d.Add "btn.reset", "Reimposta"
    d.Add "tab.scan", "Scansione"
    d.Add "tab.settings", "Impostazioni"
    d.Add "tab.log", "Registro"
    d.Add "settings.lang", "Lingua:"
    d.Add "settings.lang.auto", "Auto (host)"
    d.Add "settings.loglevel", "Livello log:"
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", "Avviso"
    d.Add "settings.loglevel.error", "Solo errori"
    d.Add "lbl.folder", "Cartella sorgente:"
    d.Add "lbl.subfolders", "Includi sottocartelle"
    d.Add "chk.force", "Forza ricreazione"
    d.Add "chk.audit", U("Verifica disegni")
    d.Add "lbl.options", "Opzioni di output:"
    d.Add "opt.single", "PDF unico (combinato)"
    d.Add "opt.multi", "File individuali"
    d.Add "chk.open", "Apri cartella al termine"
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", "Chiaro"
    d.Add "theme.dark", "Scuro"
    d.Add "theme.system", "Sistema"
    d.Add "msg.select.folder", "Seleziona una cartella da analizzare."
    d.Add "msg.no.idw", "Nessun disegno Inventor (.idw) trovato nel percorso selezionato."
    d.Add "msg.confirm.start", "Avviare l'elaborazione di {0} file?"
    d.Add "msg.process.done", "Processo completato con successo."
    d.Add "msg.error.title", "Errore di sistema"
    d.Add "log.start", "Sessione avviata: {0}"
    d.Add "log.end", "Sessione terminata: {0}"
    d.Add "log.save", "Registro salvato in: {0}"
    d.Add "scan.sum.drawings", "{0} disegni"
    d.Add "scan.sum.subfolders", "sottocartelle incluse"
    d.Add "scan.sum.ok", "{0} aggiornati"
    d.Add "scan.sum.all.current", "tutti aggiornati"
    d.Add "scan.sum.no.idw", "Nessun disegno trovato"
    d.Add "scan.sum.missing", "{0} mancanti"
    d.Add "scan.sum.stale", "{0} obsoleti"
    d.Add "scan.sum.mismatch", "{0} discordanti"
    d.Add "scan.sum.to_create", "{0} da creare"
    d.Add "scan.sum.to_recreate", "{0} da ricreare"
    d.Add "scan.sum.force_suffix", "(forzato)"
    d.Add "scan.validate.error", "Errore durante la convalida della cartella."
    d.Add "scan.sum.errors", "{0} errori"
    d.Add "scan.no.folder", "Nessuna cartella selezionata."
    d.Add "scan.no.folder.exists", "La cartella non esiste."
    d.Add "scan.path.ok", U("Cartella trovata - scansione in corso...")
    d.Add "scan.path.too.long", "Il percorso supera il limite di {0} caratteri."
    d.Add "progress.processing", "Elaborazione in corso..."
    d.Add "progress.init", "Inizializzazione..."
    d.Add "progress.cancel", "Annulla"
    d.Add "progress.cancelling", "Annullamento..."
    d.Add "scan.models.detail", "Con disegno: {0} | Senza: {1}"
    d.Add "status.idle", "Pronto"
    d.Add "status.scanning", "Scansione cartella..."
    d.Add "status.publishing", "Pubblicazione: {0} di {1}..."
    d.Add "pdf.err.open", "Impossibile aprire il file IDW."
    d.Add "pdf.err.pagematch", "Il PDF ha {0} pagine, previste {1}."
    d.Add "pdf.err.noext", "Percorso file non valido (senza estensione): {0}"
    d.Add "pdf.skipped.exists", "Ignorato (esistente)"
    d.Add "pdf.skipped.exists.pages", U("Ignorato (esistente) - {0}/{1} pagine")
    d.Add "pdf.err.addin", "AddIn PDF Translator non trovato."
    d.Add "pdf.err.nofile", "L'esportazione PDF non ha prodotto alcun file."
    d.Add "pdf.err.export", "Esportazione fallita: {0}"
    d.Add "pdf.err.bak.delete", "Impossibile rimuovere il backup obsoleto: {0}"
    d.Add "pdf.err.bak.move", "Impossibile eseguire il backup del PDF esistente: {0}"
    d.Add "pdf.err.write", "Impossibile scrivere il PDF: {0}"
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 Report di esportazione")
    d.Add "report.hdr.mode", U("Modalit\u00e0: {0}")
    d.Add "report.hdr.folder", "Cartella: {0}"
    d.Add "report.hdr.subfolders", "Sottocartelle"
    d.Add "report.hdr.force", "Sovrascrivi forzato"
    d.Add "report.hdr.generated", "Generato: {0}"
    d.Add "report.state.on", U("S\u00ec")
    d.Add "report.state.off", "No"
    d.Add "report.hdr.pdf.section", "ESPORTAZIONE PDF"
    d.Add "report.hdr.pdf.detail", "Dettagli:"
    d.Add "report.created", "Creato"
    d.Add "report.overwritten", "Sovrascritto"
    d.Add "report.skipped", "Ignorato"
    d.Add "report.errors", "Errori"
    d.Add "report.time", "Tempo"
    d.Add "report.empty", "vuoto"
    d.Add "report.pages", "pagine"
    d.Add "report.no.views", "Nessuna vista"
    d.Add "report.tag.created", "[Creato]"
    d.Add "report.tag.overwritten", "[Sovrascritto]"
    d.Add "report.tag.skipped", "[Ignorato]"
    d.Add "report.tag.error", "[Errore]"
    d.Add "report.tag.hasdrawing", "[Ha disegno]"
    d.Add "report.tag.nodrawing", "[Nessun disegno]"
    d.Add "report.hdr.audit.section", "AUDIT DISEGNI"
    d.Add "report.hdr.audit.detail", "Dettagli:"
    d.Add "report.audit.title", "Audit disegni"
    d.Add "errdlg.retry.prompt", U("Riprova = riprova | Ignora = salta | Interrompi = ferma")
    d.Add "tip.run", "Avvia esportazione PDF in batch"
    d.Add "tip.theme", "Scorri i temi (Inventor / Sistema / Chiaro / Scuro)"
    d.Add "tip.reset", "Ripristina tutte le opzioni predefinite"
    d.Add "tip.cancel", "Annulla l'operazione corrente (ESC)"
    d.Add "tip.filter.nodrawing", "Mostra solo i modelli senza disegno"
    d.Add "tip.list", "Doppio clic per aprire il PDF; clic per i dettagli"
    d.Add "tip.subfolders", "Scansione ricorsiva delle sottocartelle (salta OldVersions)"
    d.Add "tip.force", U("Ricrea il PDF anche se pi\u00f9 recente dell'IDW")
    d.Add "tip.audit", "Verifica modelli (.ipt/.iam) per disegni mancanti"
    d.Add "tip.retry", "Riprova tutti i file in errore"
    d.Add "tip.export", "Salva il report e aprilo in Notepad"
    d.Add "report.msg.close_blocked", "Impossibile chiudere durante l'elaborazione."
    d.Add "report.msg.open_failed", "Impossibile aprire: file non trovato o azione bloccata."
    d.Add "report.msg.select_row", U("Seleziona una riga per i dettagli  \u00b7  Doppio clic per aprire il PDF")
    d.Add "report.msg.retry.none", "Riprova: nessun file in stato di errore."
    d.Add "report.msg.retry.error", "Errore di riprova: {0}"
    d.Add "report.msg.retry.partial", "Riprova: {0} riusciti, {1} ancora in errore."
    d.Add "report.msg.retry.all_fixed", "Riprova: {0} riusciti, tutti risolti."
    d.Add "report.msg.copied", "Copiato!"
    d.Add "report.msg.exported", "Report aperto in Notepad."
    d.Add "report.msg.export_failed", "Impossibile salvare il report."
    d.Add "reset.msg.all_defaults", "Tutte le impostazioni ripristinate."
    d.Add "reset.lbl.registry", "Registro"
    d.Add "reset.lbl.sandbox", "Sandbox"
    d.Add "reset.lbl.state", "Stato"
    d.Add "reset.state.deleted", "Eliminato"
    d.Add "reset.state.reinitialized", "Reinizializzato"
    d.Add "reset.msg.failed", "Ripristino fallito"
    d.Add "reset.msg.registry_complete", "Registro ripristinato."
    d.Add "reset.msg.sandbox_complete", "Sandbox ripristinata"
    d.Add "scan.badge.tooltip", "Totale: {0} | Attuali: {1} | Mancanti: {2} | Obsoleti: {3} | Discordanti: {4} | Errori: {5}"
    d.Add "report.msg.opening", U("Apertura file IDW...")
    d.Add "report.msg.updating", U("Aggiornamento viste...")
    d.Add "report.msg.exporting", U("Esportazione PDF in corso...")
    d.Add "msg.no.write.access", "Impossibile scrivere nella cartella sandbox. Verifica i permessi:"
    d.Add "msg.pdf.temp_missing", "Il PDF temporaneo \u00e8 scomparso durante l'esportazione (possibile quarantena antivirus o blocco file): {0}"
    d.Add "msg.batch.abort.network", "Batch interrotto: {0} errori di rete consecutivi. Verifica la connessione e riprova."
    d.Add "msg.already.running", "Batch PDF Publisher \u00e8 gi\u00e0 in esecuzione. Chiudi l'altra sessione e riprova."
    d.Add "report.search.placeholder", "Cerca percorso..."
    d.Add "report.search.label", "Cerca:"
    d.Add "scan.collecting", "Raccolta file..."
    d.Add "scan.analysing", "Analisi: {0} / {1}..."
    d.Add "scan.auditing.start", "Verifica modelli..."
    d.Add "scan.auditing", "Verifica: {0} / {1}..."
    d.Add "scan.cancelled", "Annullato"
    d.Add "tip.search", "Filtra le righe del report per percorso"
    d.Add "msg.confirm.reset", "Ripristinare tutte le opzioni predefinite?"
    d.Add "scan.disk.insufficient", U("Spazio su disco insufficiente. Libero: {0}")
    d.Add "scan.title.no_folder", "Seleziona una cartella"
    d.Add "summary.hdr.models_full", "{0} modelli ({1} assiemi, {2} parti)"
    d.Add "summary.hdr.with_idw", "{0} con disegno"
    d.Add "summary.hdr.no_idw", "{0} senza disegno"
    d.Add "errdlg.label.context", "Contesto"
    d.Add "errdlg.label.time", "Ora"
    d.Add "errdlg.label.log", "Registro"
    d.Add "errdlg.label.stack", "Stack"
    d.Add "pdf.err.prepare", "Preparazione fallita: {0}"
    d.Add "pdf.err.header", "Il PDF temporaneo esportato non ha superato la convalida dell'intestazione"
    d.Add "pdf.err.unknown", "Errore sconosciuto"
    d.Add "report.filter.all", "Tutti"
    d.Add "report.filter.err", "Errori"
    d.Add "report.filter.skip", "Ignorati"
    d.Add "report.filter.ok", "OK"
    d.Add "report.filter.noidw", "Senza .idw"
    d.Add "scan.reset.sure", "Sicuro?"
    d.Add "scan.reset.done", "Fatto"
    d.Add "report.form.title", "Report di elaborazione"
    d.Add "report.hdr.sub", "Sottocartelle"
    d.Add "report.hdr.force_short", "Forza"
    d.Add "report.hdr.audit", "Verifica"
    d.Add "report.summary.full", U("{0} creati \u00b7 {1} sovrascritti \u00b7 {2} ignorati \u00b7 {3} errori \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} vuoti)")
    d.Add "scan.browse_title", "Seleziona cartella sorgente"
End Sub

Private Sub L_FillES(d As Object)
    d.Add "title", "Publicador de PDF por lotes"
    d.Add "btn.run", "EJECUTAR"
    d.Add "btn.cancel", "Cancelar"
    d.Add "btn.browse", "Examinar..."
    d.Add "btn.copy", "Copiar"
    d.Add "btn.help", "Ayuda"
    d.Add "btn.settings", "Ajustes"
    d.Add "btn.back", U("Atr\u00e1s")
    d.Add "btn.export", "Exportar"
    d.Add "btn.retry", "Reintentar"
    d.Add "btn.reset", "Restablecer"
    d.Add "tab.scan", "Escanear"
    d.Add "tab.settings", "Ajustes"
    d.Add "tab.log", "Registro"
    d.Add "settings.lang", "Idioma:"
    d.Add "settings.lang.auto", "Auto (host)"
    d.Add "settings.loglevel", "Nivel de log:"
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", "Aviso"
    d.Add "settings.loglevel.error", "Solo errores"
    d.Add "lbl.folder", "Carpeta de origen:"
    d.Add "lbl.subfolders", "Incluir subcarpetas"
    d.Add "chk.force", "Forzar recreaci\u00f3n"
    d.Add "chk.audit", U("Verificar dibujos")
    d.Add "lbl.options", "Opciones de salida:"
    d.Add "opt.single", U("PDF \u00fanico (combinado)")
    d.Add "opt.multi", "Archivos individuales"
    d.Add "chk.open", "Abrir carpeta al finalizar"
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", "Claro"
    d.Add "theme.dark", "Oscuro"
    d.Add "theme.system", "Sistema"
    d.Add "msg.select.folder", "Seleccione una carpeta para escanear."
    d.Add "msg.no.idw", "No se encontraron dibujos de Inventor (.idw) en la ruta seleccionada."
    d.Add "msg.confirm.start", U("\u00bfDesea iniciar el proceso para {0} archivos?")
    d.Add "msg.process.done", U("Proceso completado con \u00e9xito.")
    d.Add "msg.error.title", "Error del sistema"
    d.Add "log.start", U("Sesi\u00f3n iniciada: {0}")
    d.Add "log.end", U("Sesi\u00f3n finalizada: {0}")
    d.Add "log.save", "Registro guardado en: {0}"
    d.Add "scan.sum.drawings", "{0} dibujos"
    d.Add "scan.sum.subfolders", "incluyendo subcarpetas"
    d.Add "scan.sum.ok", "{0} vigentes"
    d.Add "scan.sum.all.current", "todos actuales"
    d.Add "scan.sum.no.idw", "No se encontraron dibujos"
    d.Add "scan.sum.missing", "{0} ausentes"
    d.Add "scan.sum.stale", "{0} desactualizados"
    d.Add "scan.sum.mismatch", "{0} discrepantes"
    d.Add "scan.sum.to_create", "{0} para crear"
    d.Add "scan.sum.to_recreate", "{0} para recrear"
    d.Add "scan.sum.force_suffix", "(forzado)"
    d.Add "scan.validate.error", "Error al validar la carpeta."
    d.Add "scan.sum.errors", "{0} errores"
    d.Add "scan.no.folder", "No se ha seleccionado ninguna carpeta."
    d.Add "scan.no.folder.exists", "La carpeta no existe."
    d.Add "scan.path.ok", U("Carpeta encontrada - escaneando...")
    d.Add "scan.path.too.long", U("La ruta supera el l\u00edmite de {0} caracteres.")
    d.Add "progress.processing", "Procesando..."
    d.Add "progress.init", "Inicializando..."
    d.Add "progress.cancel", "Cancelar"
    d.Add "progress.cancelling", "Cancelando..."
    d.Add "scan.models.detail", "Con dibujo: {0} | Sin: {1}"
    d.Add "status.idle", "Listo"
    d.Add "status.scanning", "Escaneando carpeta..."
    d.Add "status.publishing", "Publicando: {0} de {1}..."
    d.Add "pdf.err.open", "No se pudo abrir el archivo IDW."
    d.Add "pdf.err.pagematch", U("El PDF tiene {0} p\u00e1ginas, se esperaban {1}.")
    d.Add "pdf.err.noext", U("Ruta de archivo no v\u00e1lida (sin extensi\u00f3n): {0}")
    d.Add "pdf.skipped.exists", "Omitido (existente)"
    d.Add "pdf.skipped.exists.pages", U("Omitido (existente) - {0}/{1} p\u00e1ginas")
    d.Add "pdf.err.addin", "Complemento PDF Translator no encontrado."
    d.Add "pdf.err.nofile", U("La exportaci\u00f3n PDF no produjo ning\u00fan archivo.")
    d.Add "pdf.err.export", U("Exportaci\u00f3n fallida: {0}")
    d.Add "pdf.err.bak.delete", "No se puede eliminar la copia de seguridad obsoleta: {0}"
    d.Add "pdf.err.bak.move", "No se puede respaldar el PDF existente: {0}"
    d.Add "pdf.err.write", "No se puede escribir el PDF: {0}"
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 Informe de exportaci\u00f3n")
    d.Add "report.hdr.mode", "Modo: {0}"
    d.Add "report.hdr.folder", "Carpeta: {0}"
    d.Add "report.hdr.subfolders", "Subcarpetas"
    d.Add "report.hdr.force", "Sobrescritura forzada"
    d.Add "report.hdr.generated", "Generado: {0}"
    d.Add "report.state.on", "Activado"
    d.Add "report.state.off", "Desactivado"
    d.Add "report.hdr.pdf.section", U("EXPORTACI\u00d3N PDF")
    d.Add "report.hdr.pdf.detail", "Detalles:"
    d.Add "report.created", "Creado"
    d.Add "report.overwritten", "Sobrescrito"
    d.Add "report.skipped", "Omitido"
    d.Add "report.errors", "Errores"
    d.Add "report.time", "Tiempo"
    d.Add "report.empty", U("vac\u00edo")
    d.Add "report.pages", U("p\u00e1ginas")
    d.Add "report.no.views", "Sin vistas"
    d.Add "report.tag.created", "[Creado]"
    d.Add "report.tag.overwritten", "[Sobrescrito]"
    d.Add "report.tag.skipped", "[Omitido]"
    d.Add "report.tag.error", "[Error]"
    d.Add "report.tag.hasdrawing", "[Tiene dibujo]"
    d.Add "report.tag.nodrawing", "[Sin dibujo]"
    d.Add "report.hdr.audit.section", U("AUDITOR\u00cdA DE DIBUJOS")
    d.Add "report.hdr.audit.detail", "Detalles:"
    d.Add "report.audit.title", U("Auditor\u00eda de dibujos")
    d.Add "errdlg.retry.prompt", U("Reintentar = volver a intentar | Ignorar = omitir | Anular = detener")
    d.Add "tip.run", U("Iniciar exportaci\u00f3n PDF por lotes")
    d.Add "tip.theme", U("Recorrer temas (Inventor / Sistema / Claro / Oscuro)")
    d.Add "tip.reset", U("Restablecer todas las opciones")
    d.Add "tip.cancel", U("Cancelar la operaci\u00f3n actual (ESC)")
    d.Add "tip.filter.nodrawing", "Mostrar solo modelos sin dibujo"
    d.Add "tip.list", U("Doble clic para abrir el PDF; clic para ver detalles")
    d.Add "tip.subfolders", U("Escanear subcarpetas recursivamente (omite OldVersions)")
    d.Add "tip.force", U("Recrear el PDF aunque sea m\u00e1s reciente que el IDW")
    d.Add "tip.audit", U("Verificar modelos (.ipt/.iam) sin dibujo asociado")
    d.Add "tip.retry", "Reintentar todos los archivos con error"
    d.Add "tip.export", "Guardar el informe y abrirlo en Notepad"
    d.Add "report.msg.close_blocked", "No se puede cerrar durante el procesamiento."
    d.Add "report.msg.open_failed", U("No se puede abrir: archivo no encontrado o acci\u00f3n bloqueada.")
    d.Add "report.msg.select_row", U("Seleccione una fila para ver detalles  \u00b7  Doble clic para abrir el PDF")
    d.Add "report.msg.retry.none", "Reintentar: no hay archivos en estado de error."
    d.Add "report.msg.retry.error", "Error de reintento: {0}"
    d.Add "report.msg.retry.partial", U("Reintento: {0} exitosos, {1} a\u00fan con error.")
    d.Add "report.msg.retry.all_fixed", "Reintento: {0} exitosos, todos corregidos."
    d.Add "report.msg.copied", U("\u00a1Copiado!")
    d.Add "report.msg.exported", "Informe abierto en Notepad."
    d.Add "report.msg.export_failed", "No se pudo guardar el informe."
    d.Add "reset.msg.all_defaults", "Todos los ajustes restablecidos."
    d.Add "reset.lbl.registry", "Registro"
    d.Add "reset.lbl.sandbox", "Sandbox"
    d.Add "reset.lbl.state", "Estado"
    d.Add "reset.state.deleted", "Eliminado"
    d.Add "reset.state.reinitialized", "Reinicializado"
    d.Add "reset.msg.failed", "Error al restablecer"
    d.Add "reset.msg.registry_complete", "Registro restablecido."
    d.Add "reset.msg.sandbox_complete", "Sandbox restablecido"
    d.Add "scan.badge.tooltip", "Total: {0} | Actuales: {1} | Faltantes: {2} | Obsoletos: {3} | Discrepantes: {4} | Errores: {5}"
    d.Add "report.msg.opening", U("Abriendo archivo IDW...")
    d.Add "report.msg.updating", U("Actualizando vistas...")
    d.Add "report.msg.exporting", U("Exportando PDF...")
    d.Add "msg.no.write.access", "No se puede escribir en la carpeta sandbox. Verifique los permisos:"
    d.Add "msg.pdf.temp_missing", "El PDF temporal desapareci\u00f3 durante la exportaci\u00f3n (posible cuarentena de antivirus o bloqueo de archivo): {0}"
    d.Add "msg.batch.abort.network", "Lote cancelado: {0} errores de red consecutivos. Verifique la conexi\u00f3n y reintente."
    d.Add "msg.already.running", "Batch PDF Publisher ya est\u00e1 en ejecuci\u00f3n. Cierre la otra sesi\u00f3n e int\u00e9ntelo de nuevo."
    d.Add "report.search.placeholder", "Buscar ruta..."
    d.Add "report.search.label", "Buscar:"
    d.Add "scan.collecting", "Recopilando archivos..."
    d.Add "scan.analysing", U("Analizando: {0} / {1}...")
    d.Add "scan.auditing.start", U("Verificando modelos...")
    d.Add "scan.auditing", U("Verificando: {0} / {1}...")
    d.Add "scan.cancelled", "Cancelado"
    d.Add "tip.search", "Filtrar filas del informe por ruta"
    d.Add "msg.confirm.reset", U("\u00bfRestablecer todas las opciones a los valores predeterminados?")
    d.Add "scan.disk.insufficient", U("Espacio en disco insuficiente. Libre: {0}")
    d.Add "scan.title.no_folder", "Seleccione una carpeta"
    d.Add "summary.hdr.models_full", "{0} modelos ({1} ensamblajes, {2} piezas)"
    d.Add "summary.hdr.with_idw", "{0} con dibujo"
    d.Add "summary.hdr.no_idw", "{0} sin dibujo"
    d.Add "errdlg.label.context", "Contexto"
    d.Add "errdlg.label.time", "Hora"
    d.Add "errdlg.label.log", "Registro"
    d.Add "errdlg.label.stack", "Pila"
    d.Add "pdf.err.prepare", U("Preparaci\u00f3n fallida: {0}")
    d.Add "pdf.err.header", U("El PDF temporal exportado no pas\u00f3 la validaci\u00f3n de encabezado")
    d.Add "pdf.err.unknown", "Error desconocido"
    d.Add "report.filter.all", "Todos"
    d.Add "report.filter.err", "Errores"
    d.Add "report.filter.skip", "Omitidos"
    d.Add "report.filter.ok", "OK"
    d.Add "report.filter.noidw", "Sin .idw"
    d.Add "scan.reset.sure", U("\u00bfSeguro?")
    d.Add "scan.reset.done", "Hecho"
    d.Add "report.form.title", "Informe de procesamiento"
    d.Add "report.hdr.sub", "Subcarpetas"
    d.Add "report.hdr.force_short", "Forzar"
    d.Add "report.hdr.audit", "Verificar"
    d.Add "report.summary.full", U("{0} creados \u00b7 {1} sobrescritos \u00b7 {2} omitidos \u00b7 {3} errores \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} vac\u00edos)")
    d.Add "scan.browse_title", "Seleccionar carpeta de origen"
End Sub

Private Sub L_FillPT(d As Object)
    d.Add "title", "Publicador de PDF em lote"
    d.Add "btn.run", "EXECUTAR"
    d.Add "btn.cancel", "Cancelar"
    d.Add "btn.browse", "Procurar..."
    d.Add "btn.copy", "Copiar"
    d.Add "btn.help", "Ajuda"
    d.Add "btn.settings", U("Configura\u00e7\u00f5es")
    d.Add "btn.back", "Voltar"
    d.Add "btn.export", "Exportar"
    d.Add "btn.retry", "Tentar novamente"
    d.Add "btn.reset", "Redefinir"
    d.Add "tab.scan", "Verificar"
    d.Add "tab.settings", "Ajustes"
    d.Add "tab.log", "Log"
    d.Add "settings.lang", "Idioma:"
    d.Add "settings.lang.auto", "Auto (host)"
    d.Add "settings.loglevel", U("N\u00edvel de log:")
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", "Aviso"
    d.Add "settings.loglevel.error", "Apenas erros"
    d.Add "lbl.folder", "Pasta de origem:"
    d.Add "lbl.subfolders", "Incluir subpastas"
    d.Add "chk.force", "For\u00e7ar recria\u00e7\u00e3o"
    d.Add "chk.audit", U("Verificar desenhos")
    d.Add "lbl.options", U("Op\u00e7\u00f5es de sa\u00edda:")
    d.Add "opt.single", U("PDF \u00fanico (combinado)")
    d.Add "opt.multi", "Arquivos individuais"
    d.Add "chk.open", "Abrir pasta ao terminar"
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", "Claro"
    d.Add "theme.dark", "Escuro"
    d.Add "theme.system", "Sistema"
    d.Add "msg.select.folder", "Selecione uma pasta para verificar."
    d.Add "msg.no.idw", "Nenhum desenho do Inventor (.idw) encontrado."
    d.Add "msg.confirm.start", "Deseja iniciar o processo para {0} arquivos?"
    d.Add "msg.process.done", U("Processo conclu\u00eddo com sucesso.")
    d.Add "msg.error.title", "Erro do sistema"
    d.Add "log.start", U("Sess\u00e3o iniciada: {0}")
    d.Add "log.end", U("Sess\u00e3o encerrada: {0}")
    d.Add "log.save", "Log salvo em: {0}"
    d.Add "scan.sum.drawings", "{0} desenhos"
    d.Add "scan.sum.subfolders", "incl. subpastas"
    d.Add "scan.sum.ok", "{0} atualizados"
    d.Add "scan.sum.all.current", "todos atuais"
    d.Add "scan.sum.no.idw", "Nenhum desenho encontrado"
    d.Add "scan.sum.missing", "{0} ausentes"
    d.Add "scan.sum.stale", "{0} obsoletos"
    d.Add "scan.sum.mismatch", "{0} divergentes"
    d.Add "scan.sum.to_create", "{0} para criar"
    d.Add "scan.sum.to_recreate", "{0} para recriar"
    d.Add "scan.sum.force_suffix", U("(for\u00e7ado)")
    d.Add "scan.validate.error", "Erro ao validar pasta."
    d.Add "scan.sum.errors", "{0} erros"
    d.Add "scan.no.folder", "Nenhuma pasta selecionada."
    d.Add "scan.no.folder.exists", U("A pasta n\u00e3o existe.")
    d.Add "scan.path.ok", U("Pasta encontrada - verificando...")
    d.Add "scan.path.too.long", "O caminho excede o limite de {0} caracteres."
    d.Add "progress.processing", "Processando..."
    d.Add "progress.init", "Inicializando..."
    d.Add "progress.cancel", "Cancelar"
    d.Add "progress.cancelling", "Cancelando..."
    d.Add "scan.models.detail", "Com desenho: {0} | Sem: {1}"
    d.Add "status.idle", "Pronto"
    d.Add "status.scanning", "Verificando pasta..."
    d.Add "status.publishing", "Publicando: {0} de {1}..."
    d.Add "pdf.err.open", U("N\u00e3o foi poss\u00edvel abrir o arquivo IDW.")
    d.Add "pdf.err.pagematch", U("O PDF tem {0} p\u00e1ginas, esperadas {1}.")
    d.Add "pdf.err.noext", U("Caminho de arquivo inv\u00e1lido (sem extens\u00e3o): {0}")
    d.Add "pdf.skipped.exists", "Ignorado (existente)"
    d.Add "pdf.skipped.exists.pages", U("Ignorado (existente) - {0}/{1} p\u00e1ginas")
    d.Add "pdf.err.addin", U("AddIn PDF Translator n\u00e3o encontrado.")
    d.Add "pdf.err.nofile", U("A exporta\u00e7\u00e3o PDF n\u00e3o produziu nenhum arquivo.")
    d.Add "pdf.err.export", U("Exporta\u00e7\u00e3o falhou: {0}")
    d.Add "pdf.err.bak.delete", U("N\u00e3o \u00e9 poss\u00edvel remover backup obsoleto: {0}")
    d.Add "pdf.err.bak.move", U("N\u00e3o \u00e9 poss\u00edvel fazer backup do PDF existente: {0}")
    d.Add "pdf.err.write", U("N\u00e3o \u00e9 poss\u00edvel gravar o PDF: {0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 Relat\u00f3rio de exporta\u00e7\u00e3o")
    d.Add "report.hdr.mode", "Modo: {0}"
    d.Add "report.hdr.folder", "Pasta: {0}"
    d.Add "report.hdr.subfolders", "Subpastas"
    d.Add "report.hdr.force", U("Sobrescrever for\u00e7ado")
    d.Add "report.hdr.generated", "Gerado: {0}"
    d.Add "report.state.on", "Ligado"
    d.Add "report.state.off", "Desligado"
    d.Add "report.hdr.pdf.section", U("EXPORTA\u00c7\u00c3O PDF")
    d.Add "report.hdr.pdf.detail", "Detalhes:"
    d.Add "report.created", "Criado"
    d.Add "report.overwritten", "Sobrescrito"
    d.Add "report.skipped", "Ignorado"
    d.Add "report.errors", "Erros"
    d.Add "report.time", "Tempo"
    d.Add "report.empty", "vazio"
    d.Add "report.pages", U("p\u00e1ginas")
    d.Add "report.no.views", "Sem vistas"
    d.Add "report.tag.created", "[Criado]"
    d.Add "report.tag.overwritten", "[Sobrescrito]"
    d.Add "report.tag.skipped", "[Ignorado]"
    d.Add "report.tag.error", "[Erro]"
    d.Add "report.tag.hasdrawing", "[Tem desenho]"
    d.Add "report.tag.nodrawing", "[Sem desenho]"
    d.Add "report.hdr.audit.section", "AUDITORIA DE DESENHOS"
    d.Add "report.hdr.audit.detail", "Detalhes:"
    d.Add "report.audit.title", "Auditoria de desenhos"
    d.Add "errdlg.retry.prompt", U("Tentar = tentar novamente | Ignorar = ignorar | Anular = parar")
    d.Add "tip.run", U("Iniciar exporta\u00e7\u00e3o PDF em lote")
    d.Add "tip.theme", U("Percorrer temas (Inventor / Sistema / Claro / Escuro)")
    d.Add "tip.reset", U("Redefinir todas as op\u00e7\u00f5es para o padr\u00e3o")
    d.Add "tip.cancel", "Cancelar a opera\u00e7\u00e3o atual (ESC)"
    d.Add "tip.filter.nodrawing", "Mostrar apenas modelos sem desenho"
    d.Add "tip.list", U("Duplo clique abre o PDF; clique para ver detalhes")
    d.Add "tip.subfolders", U("Verificar subpastas recursivamente (ignora OldVersions)")
    d.Add "tip.force", U("Recriar o PDF mesmo se for mais recente que o IDW")
    d.Add "tip.audit", U("Verificar modelos (.ipt/.iam) sem desenho associado")
    d.Add "tip.retry", "Tentar novamente todos os arquivos com erro"
    d.Add "tip.export", U("Salvar o relat\u00f3rio e abri-lo no Notepad")
    d.Add "report.msg.close_blocked", U("N\u00e3o \u00e9 poss\u00edvel fechar durante o processamento.")
    d.Add "report.msg.open_failed", U("N\u00e3o \u00e9 poss\u00edvel abrir: arquivo n\u00e3o encontrado ou a\u00e7\u00e3o bloqueada.")
    d.Add "report.msg.select_row", U("Selecione uma linha para detalhes  \u00b7  Duplo clique para abrir o PDF")
    d.Add "report.msg.retry.none", U("Tentar novamente: nenhum arquivo em estado de erro.")
    d.Add "report.msg.retry.error", U("Erro ao tentar novamente: {0}")
    d.Add "report.msg.retry.partial", U("Tentar novamente: {0} bem-sucedidos, {1} ainda com erro.")
    d.Add "report.msg.retry.all_fixed", U("Tentar novamente: {0} bem-sucedidos, todos corrigidos.")
    d.Add "report.msg.copied", "Copiado!"
    d.Add "report.msg.exported", U("Relat\u00f3rio aberto no Notepad.")
    d.Add "report.msg.export_failed", U("Falha ao salvar o relat\u00f3rio.")
    d.Add "reset.msg.all_defaults", "Todas as configuracoes restauradas."
    d.Add "reset.lbl.registry", "Registro"
    d.Add "reset.lbl.sandbox", "Sandbox"
    d.Add "reset.lbl.state", "Estado"
    d.Add "reset.state.deleted", "Excluido"
    d.Add "reset.state.reinitialized", "Reinicializado"
    d.Add "reset.msg.failed", "Falha ao redefinir"
    d.Add "reset.msg.registry_complete", "Registro redefinido."
    d.Add "reset.msg.sandbox_complete", "Sandbox redefinido"
    d.Add "scan.badge.tooltip", "Total: {0} | Atuais: {1} | Ausentes: {2} | Obsoletos: {3} | Divergentes: {4} | Erros: {5}"
    d.Add "report.msg.opening", U("Abrindo arquivo IDW...")
    d.Add "report.msg.updating", U("Atualizando vistas...")
    d.Add "report.msg.exporting", U("Exportando PDF...")
    d.Add "msg.no.write.access", U("N\u00e3o \u00e9 poss\u00edvel gravar na pasta sandbox. Verifique as permiss\u00f5es:")
    d.Add "msg.pdf.temp_missing", U("O PDF tempor\u00e1rio desapareceu durante a exporta\u00e7\u00e3o (poss\u00edvel quarentena de antiv\u00edrus ou bloqueio de arquivo): {0}")
    d.Add "msg.batch.abort.network", U("Lote abortado: {0} falhas de rede consecutivas. Verifique a conex\u00e3o e tente novamente.")
    d.Add "msg.already.running", U("O Batch PDF Publisher j\u00e1 est\u00e1 em execu\u00e7\u00e3o. Feche a outra sess\u00e3o e tente novamente.")
    d.Add "report.search.placeholder", "Pesquisar caminho..."
    d.Add "report.search.label", "Pesquisar:"
    d.Add "scan.collecting", "Coletando arquivos..."
    d.Add "scan.analysing", "Analisando: {0} / {1}..."
    d.Add "scan.auditing.start", "Verificando modelos..."
    d.Add "scan.auditing", "Verificando: {0} / {1}..."
    d.Add "scan.cancelled", "Cancelado"
    d.Add "tip.search", "Filtrar linhas do relat\u00f3rio por caminho"
    d.Add "msg.confirm.reset", U("Redefinir todas as op\u00e7\u00f5es para o padr\u00e3o?")
    d.Add "scan.disk.insufficient", U("Espa\u00e7o em disco insuficiente. Livre: {0}")
    d.Add "scan.title.no_folder", "Selecione uma pasta"
    d.Add "summary.hdr.models_full", "{0} modelos ({1} conjuntos, {2} pe\u00e7as)"
    d.Add "summary.hdr.with_idw", "{0} com desenho"
    d.Add "summary.hdr.no_idw", "{0} sem desenho"
    d.Add "errdlg.label.context", "Contexto"
    d.Add "errdlg.label.time", "Hora"
    d.Add "errdlg.label.log", "Log"
    d.Add "errdlg.label.stack", "Pilha"
    d.Add "pdf.err.prepare", U("Prepara\u00e7\u00e3o falhou: {0}")
    d.Add "pdf.err.header", U("O PDF tempor\u00e1rio exportado falhou na valida\u00e7\u00e3o de cabe\u00e7alho")
    d.Add "pdf.err.unknown", "Erro desconhecido"
    d.Add "report.filter.all", "Todos"
    d.Add "report.filter.err", "Erros"
    d.Add "report.filter.skip", "Ignorados"
    d.Add "report.filter.ok", "OK"
    d.Add "report.filter.noidw", "Sem .idw"
    d.Add "scan.reset.sure", "Certo?"
    d.Add "scan.reset.done", "Feito"
    d.Add "report.form.title", "Relat\u00f3rio de processamento"
    d.Add "report.hdr.sub", "Subpastas"
    d.Add "report.hdr.force_short", "For\u00e7ar"
    d.Add "report.hdr.audit", "Verificar"
    d.Add "report.summary.full", U("{0} criados \u00b7 {1} sobrescritos \u00b7 {2} ignorados \u00b7 {3} erros \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} vazios)")
    d.Add "scan.browse_title", "Selecionar pasta de origem"
End Sub

Private Sub L_FillRU(d As Object)
    d.Add "title", U("\u041f\u0430\u043a\u0435\u0442\u043d\u0430\u044f \u043f\u0443\u0431\u043b\u0438\u043a\u0430\u0446\u0438\u044f PDF")
    d.Add "btn.run", U("\u0417\u0410\u041f\u0423\u0421\u041a")
    d.Add "btn.cancel", U("\u041e\u0442\u043c\u0435\u043d\u0430")
    d.Add "btn.browse", U("\u041e\u0431\u0437\u043e\u0440...")
    d.Add "btn.copy", U("\u041a\u043e\u043f\u0438\u0440\u043e\u0432\u0430\u0442\u044c")
    d.Add "btn.help", U("\u0421\u043f\u0440\u0430\u0432\u043a\u0430")
    d.Add "btn.settings", U("\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438")
    d.Add "btn.back", U("\u041d\u0430\u0437\u0430\u0434")
    d.Add "btn.export", U("\u042d\u043a\u0441\u043f\u043e\u0440\u0442")
    d.Add "btn.retry", U("\u041f\u043e\u0432\u0442\u043e\u0440\u0438\u0442\u044c")
    d.Add "btn.reset", U("\u0421\u0431\u0440\u043e\u0441")
    d.Add "tab.scan", U("\u041f\u043e\u0438\u0441\u043a")
    d.Add "tab.settings", U("\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438")
    d.Add "tab.log", U("\u0416\u0443\u0440\u043d\u0430\u043b")
    d.Add "settings.lang", U("\u042f\u0437\u044b\u043a:")
    d.Add "settings.lang.auto", U("\u0410\u0432\u0442\u043e (\u0445\u043e\u0441\u0442)")
    d.Add "settings.loglevel", U("\u0423\u0440\u043e\u0432\u0435\u043d\u044c \u0436\u0443\u0440\u043d\u0430\u043b\u0430:")
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", U("\u041f\u0440\u0435\u0434\u0443\u043f\u0440\u0435\u0436\u0434\u0435\u043d\u0438\u0435")
    d.Add "settings.loglevel.error", U("\u0422\u043e\u043b\u044c\u043a\u043e \u043e\u0448\u0438\u0431\u043a\u0438")
    d.Add "lbl.folder", U("\u0418\u0441\u0445\u043e\u0434\u043d\u0430\u044f \u043f\u0430\u043f\u043a\u0430:")
    d.Add "lbl.subfolders", U("\u0412\u043a\u043b\u044e\u0447\u0430\u044f \u043f\u043e\u0434\u043f\u0430\u043f\u043a\u0438")
    d.Add "chk.force", U("\u041f\u0440\u0438\u043d\u0443\u0434\u0438\u0442\u0435\u043b\u044c\u043d\u043e \u043f\u0435\u0440\u0435\u0441\u043e\u0437\u0434\u0430\u0442\u044c")
    d.Add "chk.audit", U("\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u0447\u0435\u0440\u0442\u0435\u0436\u0435\u0439")
    d.Add "lbl.options", U("\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b \u0432\u044b\u0432\u043e\u0434\u0430:")
    d.Add "opt.single", U("\u0415\u0434\u0438\u043d\u044b\u0439 PDF (\u0441\u0431\u043e\u0440\u043d\u044b\u0439)")
    d.Add "opt.multi", U("\u041e\u0442\u0434\u0435\u043b\u044c\u043d\u044b\u0435 \u0444\u0430\u0439\u043b\u044b")
    d.Add "chk.open", U("\u041e\u0442\u043a\u0440\u044b\u0442\u044c \u043f\u0430\u043f\u043a\u0443 \u043f\u043e \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u0438\u0438")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", U("\u0421\u0432\u0435\u0442\u043b\u0430\u044f")
    d.Add "theme.dark", U("\u0422\u0435\u043c\u043d\u0430\u044f")
    d.Add "theme.system", U("\u0421\u0438\u0441\u0442\u0435\u043c\u043d\u0430\u044f")
    d.Add "msg.select.folder", U("\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043f\u0430\u043f\u043a\u0443 \u0434\u043b\u044f \u0441\u043a\u0430\u043d\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u044f.")
    d.Add "msg.no.idw", U("\u0427\u0435\u0440\u0442\u0435\u0436\u0438 Inventor (.idw) \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b.")
    d.Add "msg.confirm.start", U("\u041d\u0430\u0447\u0430\u0442\u044c \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0443 {0} \u0444\u0430\u0439\u043b\u043e\u0432?")
    d.Add "msg.process.done", U("\u041f\u0440\u043e\u0446\u0435\u0441\u0441 \u0443\u0441\u043f\u0435\u0448\u043d\u043e \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d.")
    d.Add "msg.error.title", U("\u0421\u0438\u0441\u0442\u0435\u043c\u043d\u0430\u044f \u043e\u0448\u0438\u0431\u043a\u0430")
    d.Add "log.start", U("\u0421\u0435\u0441\u0441\u0438\u044f \u043d\u0430\u0447\u0430\u0442\u0430: {0}")
    d.Add "log.end", U("\u0421\u0435\u0441\u0441\u0438\u044f \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u0430: {0}")
    d.Add "log.save", U("\u0416\u0443\u0440\u043d\u0430\u043b \u0441\u043e\u0445\u0440\u0430\u043d\u0435\u043d: {0}")
    d.Add "scan.sum.drawings", U("{0} \u0447\u0435\u0440\u0442\u0435\u0436\u0435\u0439")
    d.Add "scan.sum.subfolders", U("\u0441 \u043f\u043e\u0434\u043f\u0430\u043f\u043a\u0430\u043c\u0438")
    d.Add "scan.sum.ok", U("{0} \u0430\u043a\u0442\u0443\u0430\u043b\u044c\u043d\u043e")
    d.Add "scan.sum.all.current", U("\u0432\u0441\u0435 \u0430\u043a\u0442\u0443\u0430\u043b\u044c\u043d\u044b")
    d.Add "scan.sum.no.idw", U("\u0427\u0435\u0440\u0442\u0435\u0436\u0438 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b")
    d.Add "scan.sum.missing", U("{0} \u043e\u0442\u0441\u0443\u0442\u0441\u0442\u0432\u0443\u044e\u0442")
    d.Add "scan.sum.stale", U("{0} \u0443\u0441\u0442\u0430\u0440\u0435\u043b\u043e")
    d.Add "scan.sum.mismatch", U("{0} \u043d\u0435\u0441\u043e\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0438\u0439")
    d.Add "scan.sum.to_create", U("{0} \u0441\u043e\u0437\u0434\u0430\u0442\u044c")
    d.Add "scan.sum.to_recreate", U("{0} \u043f\u0435\u0440\u0435\u0441\u043e\u0437\u0434\u0430\u0442\u044c")
    d.Add "scan.sum.force_suffix", U("(\u043f\u0440\u0438\u043d\u0443\u0434\u0438\u0442\u0435\u043b\u044c\u043d\u043e)")
    d.Add "scan.validate.error", U("\u041e\u0448\u0438\u0431\u043a\u0430 \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438 \u043f\u0430\u043f\u043a\u0438.")
    d.Add "scan.sum.errors", U("{0} \u043e\u0448\u0438\u0431\u043e\u043a")
    d.Add "scan.no.folder", U("\u041f\u0430\u043f\u043a\u0430 \u043d\u0435 \u0432\u044b\u0431\u0440\u0430\u043d\u0430.")
    d.Add "scan.no.folder.exists", U("\u041f\u0430\u043f\u043a\u0430 \u043d\u0435 \u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u0435\u0442.")
    d.Add "scan.path.ok", U("\u041f\u0430\u043f\u043a\u0430 \u043d\u0430\u0439\u0434\u0435\u043d\u0430 \u2014 \u0441\u043a\u0430\u043d\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u0435...")
    d.Add "scan.path.too.long", U("\u041f\u0443\u0442\u044c \u043f\u0440\u0435\u0432\u044b\u0448\u0430\u0435\u0442 \u043b\u0438\u043c\u0438\u0442 {0} \u0441\u0438\u043c\u0432\u043e\u043b\u043e\u0432.")
    d.Add "progress.processing", U("\u041e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0430...")
    d.Add "progress.init", U("\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u0438\u0437\u0430\u0446\u0438\u044f...")
    d.Add "progress.cancel", U("\u041e\u0442\u043c\u0435\u043d\u0430")
    d.Add "progress.cancelling", U("\u041e\u0442\u043c\u0435\u043d\u0430...")
    d.Add "scan.models.detail", U("\u0421 \u0447\u0435\u0440\u0442\u0435\u0436\u043e\u043c: {0} | \u0411\u0435\u0437: {1}")
    d.Add "status.idle", U("\u0413\u043e\u0442\u043e\u0432")
    d.Add "status.scanning", U("\u0421\u043a\u0430\u043d\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u0435...")
    d.Add "status.publishing", U("\u041f\u0443\u0431\u043b\u0438\u043a\u0430\u0446\u0438\u044f: {0} \u0438\u0437 {1}...")
    d.Add "pdf.err.open", U("\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043e\u0442\u043a\u0440\u044b\u0442\u044c \u0444\u0430\u0439\u043b IDW.")
    d.Add "pdf.err.pagematch", U("PDF \u0441\u043e\u0434\u0435\u0440\u0436\u0438\u0442 {0} \u0441\u0442\u0440., \u043e\u0436\u0438\u0434\u0430\u043b\u043e\u0441\u044c {1}.")
    d.Add "pdf.err.noext", U("\u041d\u0435\u0432\u0435\u0440\u043d\u044b\u0439 \u043f\u0443\u0442\u044c (\u043d\u0435\u0442 \u0440\u0430\u0441\u0448\u0438\u0440\u0435\u043d\u0438\u044f): {0}")
    d.Add "pdf.skipped.exists", U("\u041f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u043e (\u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u0435\u0442)")
    d.Add "pdf.skipped.exists.pages", U("\u041f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u043e (\u0441\u0443\u0449.) \u00b7 {0}/{1} \u0441\u0442\u0440.")
    d.Add "pdf.err.addin", U("\u0414\u043e\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u0435 PDF Translator \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u043e.")
    d.Add "pdf.err.nofile", U("\u042d\u043a\u0441\u043f\u043e\u0440\u0442 PDF \u043d\u0435 \u0441\u043e\u0437\u0434\u0430\u043b \u0444\u0430\u0439\u043b.")
    d.Add "pdf.err.export", U("\u042d\u043a\u0441\u043f\u043e\u0440\u0442 \u043d\u0435 \u0443\u0434\u0430\u043b\u0441\u044f: {0}")
    d.Add "pdf.err.bak.delete", U("\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0443\u0434\u0430\u043b\u0438\u0442\u044c \u0441\u0442\u0430\u0440\u044b\u0439 \u0431\u044d\u043a\u0430\u043f: {0}")
    d.Add "pdf.err.bak.move", U("\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0441\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c \u043a\u043e\u043f\u0438\u044e PDF: {0}")
    d.Add "pdf.err.write", U("\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0437\u0430\u043f\u0438\u0441\u0430\u0442\u044c PDF: {0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 \u041e\u0442\u0447\u0451\u0442 \u044d\u043a\u0441\u043f\u043e\u0440\u0442\u0430")
    d.Add "report.hdr.mode", U("\u0420\u0435\u0436\u0438\u043c: {0}")
    d.Add "report.hdr.folder", U("\u041f\u0430\u043f\u043a\u0430: {0}")
    d.Add "report.hdr.subfolders", U("\u041f\u043e\u0434\u043f\u0430\u043f\u043a\u0438")
    d.Add "report.hdr.force", U("\u041f\u0440\u0438\u043d\u0443\u0434\u0438\u0442\u0435\u043b\u044c\u043d\u0430\u044f \u0437\u0430\u043f\u0438\u0441\u044c")
    d.Add "report.hdr.generated", U("\u0421\u0444\u043e\u0440\u043c\u0438\u0440\u043e\u0432\u0430\u043d: {0}")
    d.Add "report.state.on", U("\u0412\u043a\u043b.")
    d.Add "report.state.off", U("\u0412\u044b\u043a\u043b.")
    d.Add "report.hdr.pdf.section", U("\u042d\u041a\u0421\u041f\u041e\u0420\u0422 PDF")
    d.Add "report.hdr.pdf.detail", U("\u0414\u0435\u0442\u0430\u043b\u0438:")
    d.Add "report.created", U("\u0421\u043e\u0437\u0434\u0430\u043d")
    d.Add "report.overwritten", U("\u041f\u0435\u0440\u0435\u0437\u0430\u043f\u0438\u0441\u0430\u043d")
    d.Add "report.skipped", U("\u041f\u0440\u043e\u043f\u0443\u0449\u0435\u043d")
    d.Add "report.errors", U("\u041e\u0448\u0438\u0431\u043a\u0438")
    d.Add "report.time", U("\u0412\u0440\u0435\u043c\u044f")
    d.Add "report.empty", U("\u043f\u0443\u0441\u0442\u043e\u0439")
    d.Add "report.pages", U("\u0441\u0442\u0440.")
    d.Add "report.no.views", U("\u041d\u0435\u0442 \u0432\u0438\u0434\u043e\u0432")
    d.Add "report.tag.created", U("[\u0421\u043e\u0437\u0434\u0430\u043d]")
    d.Add "report.tag.overwritten", U("[\u041f\u0435\u0440\u0435\u0437\u0430\u043f\u0438\u0441\u0430\u043d]")
    d.Add "report.tag.skipped", U("[\u041f\u0440\u043e\u043f\u0443\u0449\u0435\u043d]")
    d.Add "report.tag.error", U("[\u041e\u0448\u0438\u0431\u043a\u0430]")
    d.Add "report.tag.hasdrawing", U("[\u0415\u0441\u0442\u044c \u0447\u0435\u0440\u0442\u0451\u0436]")
    d.Add "report.tag.nodrawing", U("[\u041d\u0435\u0442 \u0447\u0435\u0440\u0442\u0435\u0436\u0430]")
    d.Add "report.hdr.audit.section", U("\u0410\u0423\u0414\u0418\u0422 \u0427\u0415\u0420\u0422\u0415\u0416\u0415\u0419")
    d.Add "report.hdr.audit.detail", U("\u0414\u0435\u0442\u0430\u043b\u0438:")
    d.Add "report.audit.title", U("\u0410\u0443\u0434\u0438\u0442 \u0447\u0435\u0440\u0442\u0435\u0436\u0435\u0439")
    d.Add "errdlg.retry.prompt", U("\u041f\u043e\u0432\u0442\u043e\u0440 = \u043f\u043e\u0432\u0442\u043e\u0440\u0438\u0442\u044c | \u041f\u0440\u043e\u043f\u0443\u0441\u0442\u0438\u0442\u044c = \u043f\u0440\u043e\u043f\u0443\u0441\u0442\u0438\u0442\u044c | \u041f\u0440\u0435\u0440\u0432\u0430\u0442\u044c = \u043e\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c")
    d.Add "tip.run", U("\u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u044c \u043f\u0430\u043a\u0435\u0442\u043d\u044b\u0439 \u044d\u043a\u0441\u043f\u043e\u0440\u0442 PDF")
    d.Add "tip.theme", U("\u041f\u0435\u0440\u0435\u043a\u043b\u044e\u0447\u0430\u0442\u044c \u0442\u0435\u043c\u044b (Inventor / \u0421\u0438\u0441\u0442\u0435\u043c\u043d\u0430\u044f / \u0421\u0432\u0435\u0442\u043b\u0430\u044f / \u0422\u0435\u043c\u043d\u0430\u044f)")
    d.Add "tip.reset", U("\u0421\u0431\u0440\u043e\u0441\u0438\u0442\u044c \u0432\u0441\u0435 \u043d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438 \u043f\u043e \u0443\u043c\u043e\u043b\u0447\u0430\u043d\u0438\u044e")
    d.Add "tip.cancel", U("\u041e\u0442\u043c\u0435\u043d\u0438\u0442\u044c \u0442\u0435\u043a\u0443\u0449\u0443\u044e \u043e\u043f\u0435\u0440\u0430\u0446\u0438\u044e (ESC)")
    d.Add "tip.filter.nodrawing", U("\u041f\u043e\u043a\u0430\u0437\u0430\u0442\u044c \u0442\u043e\u043b\u044c\u043a\u043e \u043c\u043e\u0434\u0435\u043b\u0438 \u0431\u0435\u0437 \u0447\u0435\u0440\u0442\u0435\u0436\u0430")
    d.Add "tip.list", U("\u0414\u0432\u043e\u0439\u043d\u043e\u0439 \u043a\u043b\u0438\u043a \u043e\u0442\u043a\u0440\u044b\u0432\u0430\u0435\u0442 PDF; \u043a\u043b\u0438\u043a \u043f\u043e\u043a\u0430\u0437\u044b\u0432\u0430\u0435\u0442 \u0434\u0435\u0442\u0430\u043b\u0438")
    d.Add "tip.subfolders", U("\u0420\u0435\u043a\u0443\u0440\u0441\u0438\u0432\u043d\u043e \u0441\u043a\u0430\u043d\u0438\u0440\u043e\u0432\u0430\u0442\u044c \u043f\u043e\u0434\u043f\u0430\u043f\u043a\u0438 (\u0438\u0441\u043a\u043b\u044e\u0447\u0430\u044f OldVersions)")
    d.Add "tip.force", U("\u041f\u0435\u0440\u0435\u0441\u043e\u0437\u0434\u0430\u0442\u044c PDF, \u0434\u0430\u0436\u0435 \u0435\u0441\u043b\u0438 \u043e\u043d \u043d\u043e\u0432\u0435\u0435 IDW")
    d.Add "tip.audit", U("\u041f\u0440\u043e\u0432\u0435\u0440\u0438\u0442\u044c \u043c\u043e\u0434\u0435\u043b\u0438 (.ipt/.iam) \u043d\u0430 \u043e\u0442\u0441\u0443\u0442\u0441\u0442\u0432\u0443\u044e\u0449\u0438\u0435 \u0447\u0435\u0440\u0442\u0435\u0436\u0438")
    d.Add "tip.retry", U("\u041f\u043e\u0432\u0442\u043e\u0440\u0438\u0442\u044c \u0432\u0441\u0435 \u0444\u0430\u0439\u043b\u044b \u0441 \u043e\u0448\u0438\u0431\u043a\u0430\u043c\u0438")
    d.Add "tip.export", U("\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c \u043e\u0442\u0447\u0451\u0442 \u0432 \u0444\u0430\u0439\u043b \u0438 \u043e\u0442\u043a\u0440\u044b\u0442\u044c \u0432 \u0411\u043b\u043e\u043a\u043d\u043e\u0442\u0435")
    d.Add "report.msg.close_blocked", U("\u041d\u0435\u0432\u043e\u0437\u043c\u043e\u0436\u043d\u043e \u0437\u0430\u043a\u0440\u044b\u0442\u044c \u0432\u043e \u0432\u0440\u0435\u043c\u044f \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0438.")
    d.Add "report.msg.open_failed", U("\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043e\u0442\u043a\u0440\u044b\u0442\u044c: \u0444\u0430\u0439\u043b \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d \u0438\u043b\u0438 \u0434\u0435\u0439\u0441\u0442\u0432\u0438\u0435 \u0437\u0430\u0431\u043b\u043e\u043a\u0438\u0440\u043e\u0432\u0430\u043d\u043e.")
    d.Add "report.msg.select_row", U("\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u0441\u0442\u0440\u043e\u043a\u0443 \u0434\u043b\u044f \u043f\u043e\u0434\u0440\u043e\u0431\u043d\u043e\u0441\u0442\u0435\u0439  \u00b7  \u0414\u0432\u043e\u0439\u043d\u043e\u0439 \u043a\u043b\u0438\u043a \u0434\u043b\u044f \u043e\u0442\u043a\u0440\u044b\u0442\u0438\u044f PDF")
    d.Add "report.msg.retry.none", U("\u041f\u043e\u0432\u0442\u043e\u0440: \u043d\u0435\u0442 \u0444\u0430\u0439\u043b\u043e\u0432 \u0441 \u043e\u0448\u0438\u0431\u043a\u0430\u043c\u0438.")
    d.Add "report.msg.retry.error", U("\u041e\u0448\u0438\u0431\u043a\u0430 \u043f\u043e\u0432\u0442\u043e\u0440\u0430: {0}")
    d.Add "report.msg.retry.partial", U("\u041f\u043e\u0432\u0442\u043e\u0440: {0} \u0443\u0441\u043f\u0435\u0448\u043d\u043e, {1} \u0432\u0441\u0451 \u0435\u0449\u0451 \u0441 \u043e\u0448\u0438\u0431\u043a\u043e\u0439.")
    d.Add "report.msg.retry.all_fixed", U("\u041f\u043e\u0432\u0442\u043e\u0440: {0} \u0443\u0441\u043f\u0435\u0448\u043d\u043e, \u0432\u0441\u0435 \u0438\u0441\u043f\u0440\u0430\u0432\u043b\u0435\u043d\u044b.")
    d.Add "report.msg.copied", U("\u0421\u043a\u043e\u043f\u0438\u0440\u043e\u0432\u0430\u043d\u043e!")
    d.Add "report.msg.exported", U("\u041e\u0442\u0447\u0451\u0442 \u043e\u0442\u043a\u0440\u044b\u0442 \u0432 \u0411\u043b\u043e\u043a\u043d\u043e\u0442\u0435.")
    d.Add "report.msg.export_failed", U("\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0441\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c \u043e\u0442\u0447\u0451\u0442.")
    d.Add "reset.msg.all_defaults", U("\u0412\u0441\u0435 \u043d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438 \u0441\u0431\u0440\u043e\u0448\u0435\u043d\u044b.")
    d.Add "reset.lbl.registry", U("\u0420\u0435\u0435\u0441\u0442\u0440")
    d.Add "reset.lbl.sandbox", "Sandbox"
    d.Add "reset.lbl.state", U("\u0421\u043e\u0441\u0442\u043e\u044f\u043d\u0438\u0435")
    d.Add "reset.state.deleted", U("\u0423\u0434\u0430\u043b\u0435\u043d\u043e")
    d.Add "reset.state.reinitialized", U("\u041f\u0435\u0440\u0435\u0438\u043d\u0438\u0446\u0438\u0430\u043b\u0438\u0437\u0438\u0440\u043e\u0432\u0430\u043d\u043e")
    d.Add "reset.msg.failed", U("\u041e\u0448\u0438\u0431\u043a\u0430 \u0441\u0431\u0440\u043e\u0441\u0430")
    d.Add "reset.msg.registry_complete", U("\u0420\u0435\u0435\u0441\u0442\u0440 \u0441\u0431\u0440\u043e\u0448\u0435\u043d.")
    d.Add "reset.msg.sandbox_complete", U("Sandbox \u0441\u0431\u0440\u043e\u0448\u0435\u043d")
    d.Add "scan.badge.tooltip", U("\u0412\u0441\u0435\u0433\u043e: {0} | \u0410\u043a\u0442\u0443\u0430\u043b\u044c\u043d\u044b\u0445: {1} | \u041e\u0442\u0441\u0443\u0442\u0441\u0442\u0432\u0443\u044e\u0442: {2} | \u0423\u0441\u0442\u0430\u0440\u0435\u043b\u043e: {3} | \u041d\u0435\u0441\u043e\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0438\u0439: {4} | \u041e\u0448\u0438\u0431\u043e\u043a: {5}")
    d.Add "report.msg.opening", U("\u041e\u0442\u043a\u0440\u044b\u0442\u0438\u0435 \u0444\u0430\u0439\u043b\u0430 IDW...")
    d.Add "report.msg.updating", U("\u041e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u0435 \u0432\u0438\u0434\u043e\u0432...")
    d.Add "report.msg.exporting", U("\u042d\u043a\u0441\u043f\u043e\u0440\u0442 PDF...")
    d.Add "msg.no.write.access", U("\u041d\u0435\u0432\u043e\u0437\u043c\u043e\u0436\u043d\u043e \u0437\u0430\u043f\u0438\u0441\u0430\u0442\u044c \u0432 \u043f\u0430\u043f\u043a\u0443 sandbox. \u041f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u043f\u0440\u0430\u0432\u0430 \u0434\u043e\u0441\u0442\u0443\u043f\u0430:")
    d.Add "msg.pdf.temp_missing", U("\u0412\u0440\u0435\u043c\u0435\u043d\u043d\u044b\u0439 PDF \u0438\u0441\u0447\u0435\u0437 \u0432\u043e \u0432\u0440\u0435\u043c\u044f \u044d\u043a\u0441\u043f\u043e\u0440\u0442\u0430 (\u0432\u043e\u0437\u043c\u043e\u0436\u043d\u043e, \u043a\u0430\u0440\u0430\u043d\u0442\u0438\u043d \u0430\u043d\u0442\u0438\u0432\u0438\u0440\u0443\u0441\u0430 \u0438\u043b\u0438 \u0431\u043b\u043e\u043a\u0438\u0440\u043e\u0432\u043a\u0430 \u0444\u0430\u0439\u043b\u0430): {0}")
    d.Add "msg.batch.abort.network", U("\u041f\u0430\u043a\u0435\u0442 \u043f\u0440\u0435\u0440\u0432\u0430\u043d: {0} \u043f\u043e\u0441\u043b\u0435\u0434\u043e\u0432\u0430\u0442\u0435\u043b\u044c\u043d\u044b\u0445 \u0441\u0435\u0442\u0435\u0432\u044b\u0445 \u043e\u0448\u0438\u0431\u043e\u043a. \u041f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0441\u043e\u0435\u0434\u0438\u043d\u0435\u043d\u0438\u0435 \u0438 \u043f\u043e\u0432\u0442\u043e\u0440\u0438\u0442\u0435.")
    d.Add "msg.already.running", U("Batch PDF Publisher \u0443\u0436\u0435 \u0437\u0430\u043f\u0443\u0449\u0435\u043d. \u0417\u0430\u043a\u0440\u043e\u0439\u0442\u0435 \u0434\u0440\u0443\u0433\u0443\u044e \u0441\u0435\u0441\u0441\u0438\u044e \u0438 \u043f\u043e\u0432\u0442\u043e\u0440\u0438\u0442\u0435.")
    d.Add "report.search.placeholder", U("\u041f\u043e\u0438\u0441\u043a \u043f\u0443\u0442\u0438...")
    d.Add "report.search.label", U("\u041f\u043e\u0438\u0441\u043a:")
    d.Add "scan.collecting", U("\u0421\u0431\u043e\u0440 \u0444\u0430\u0439\u043b\u043e\u0432...")
    d.Add "scan.analysing", U("\u0410\u043d\u0430\u043b\u0438\u0437: {0} / {1}...")
    d.Add "scan.auditing.start", U("\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u043c\u043e\u0434\u0435\u043b\u0435\u0439...")
    d.Add "scan.auditing", U("\u0410\u0443\u0434\u0438\u0442: {0} / {1}...")
    d.Add "scan.cancelled", U("\u041e\u0442\u043c\u0435\u043d\u0435\u043d\u043e")
    d.Add "tip.search", U("\u0424\u0438\u043b\u044c\u0442\u0440 \u0441\u0442\u0440\u043e\u043a \u043e\u0442\u0447\u0451\u0442\u0430 \u043f\u043e \u043f\u0443\u0442\u0438")
    d.Add "msg.confirm.reset", U("\u0421\u0431\u0440\u043e\u0441\u0438\u0442\u044c \u0432\u0441\u0435 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b \u043f\u043e \u0443\u043c\u043e\u043b\u0447\u0430\u043d\u0438\u044e?")
    d.Add "scan.disk.insufficient", U("\u041d\u0435\u0434\u043e\u0441\u0442\u0430\u0442\u043e\u0447\u043d\u043e \u043c\u0435\u0441\u0442\u0430 \u043d\u0430 \u0434\u0438\u0441\u043a\u0435. \u0421\u0432\u043e\u0431\u043e\u0434\u043d\u043e: {0}")
    d.Add "scan.title.no_folder", U("\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043f\u0430\u043f\u043a\u0443")
    d.Add "summary.hdr.models_full", U("{0} \u043c\u043e\u0434\u0435\u043b\u0435\u0439 ({1} \u0441\u0431\u043e\u0440\u043e\u043a, {2} \u0434\u0435\u0442\u0430\u043b\u0435\u0439)")
    d.Add "summary.hdr.with_idw", U("{0} \u0441 \u0447\u0435\u0440\u0442\u0435\u0436\u043e\u043c")
    d.Add "summary.hdr.no_idw", U("{0} \u0431\u0435\u0437 \u0447\u0435\u0440\u0442\u0435\u0436\u0430")
    d.Add "errdlg.label.context", U("\u041a\u043e\u043d\u0442\u0435\u043a\u0441\u0442")
    d.Add "errdlg.label.time", U("\u0412\u0440\u0435\u043c\u044f")
    d.Add "errdlg.label.log", U("\u0416\u0443\u0440\u043d\u0430\u043b")
    d.Add "errdlg.label.stack", U("\u0421\u0442\u0435\u043a")
    d.Add "pdf.err.prepare", U("\u041f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u043a\u0430 \u043d\u0435 \u0443\u0434\u0430\u043b\u0430\u0441\u044c: {0}")
    d.Add "pdf.err.header", U("\u042d\u043a\u0441\u043f\u043e\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u043d\u043d\u044b\u0439 \u0432\u0440\u0435\u043c\u0435\u043d\u043d\u044b\u0439 PDF \u043d\u0435 \u043f\u0440\u043e\u0448\u0451\u043b \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0443 \u0437\u0430\u0433\u043e\u043b\u043e\u0432\u043a\u0430")
    d.Add "pdf.err.unknown", U("\u041d\u0435\u0438\u0437\u0432\u0435\u0441\u0442\u043d\u0430\u044f \u043e\u0448\u0438\u0431\u043a\u0430")
    d.Add "report.filter.all", U("\u0412\u0441\u0435")
    d.Add "report.filter.err", U("\u041e\u0448\u0438\u0431\u043a\u0438")
    d.Add "report.filter.skip", U("\u041f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u043e")
    d.Add "report.filter.ok", U("\u041e\u041a")
    d.Add "report.filter.noidw", U("\u0411\u0435\u0437 .idw")
    d.Add "scan.reset.sure", U("\u0422\u043e\u0447\u043d\u043e?")
    d.Add "scan.reset.done", U("\u0413\u043e\u0442\u043e\u0432\u043e")
    d.Add "report.form.title", U("\u041e\u0442\u0447\u0451\u0442 \u043e\u0431 \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0435")
    d.Add "report.hdr.sub", U("\u041f\u043e\u0434\u043f\u0430\u043f\u043a\u0438")
    d.Add "report.hdr.force_short", U("\u041f\u0440\u0438\u043d\u0443\u0434.")
    d.Add "report.hdr.audit", U("\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430")
    d.Add "report.summary.full", U("{0} \u0441\u043e\u0437\u0434\u0430\u043d\u043e \u00b7 {1} \u043f\u0435\u0440\u0435\u0437\u0430\u043f\u0438\u0441\u0430\u043d\u043e \u00b7 {2} \u043f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u043e \u00b7 {3} \u043e\u0448\u0438\u0431\u043e\u043a \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} \u043f\u0443\u0441\u0442\u043e)")
    d.Add "scan.browse_title", U("\u0412\u044b\u0431\u0440\u0430\u0442\u044c \u0438\u0441\u0445\u043e\u0434\u043d\u0443\u044e \u043f\u0430\u043f\u043a\u0443")
End Sub

Private Sub L_FillPL(d As Object)
    d.Add "title", "Masowe publikowanie PDF"
    d.Add "btn.run", "URUCHOM"
    d.Add "btn.cancel", "Anuluj"
    d.Add "btn.browse", U("Przegl\u0105daj...")
    d.Add "btn.copy", "Kopiuj"
    d.Add "btn.help", "Pomoc"
    d.Add "btn.settings", "Ustawienia"
    d.Add "btn.back", "Wstecz"
    d.Add "btn.export", "Eksportuj"
    d.Add "btn.retry", U("Pon\u00f3w")
    d.Add "btn.reset", "Resetuj"
    d.Add "tab.scan", "Skanuj"
    d.Add "tab.settings", "Ustawienia"
    d.Add "tab.log", "Log"
    d.Add "settings.lang", U("J\u0119zyk:")
    d.Add "settings.lang.auto", "Auto (host)"
    d.Add "settings.loglevel", "Poziom logu:"
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", U("Ostrze\u017cenie")
    d.Add "settings.loglevel.error", U("Tylko b\u0142\u0119dy")
    d.Add "lbl.folder", U("Folder \u017ar\u00f3d\u0142owy:")
    d.Add "lbl.subfolders", U("Uwzgl\u0119dnij podfoldery")
    d.Add "chk.force", U("Wymu\u015b ponowne utworzenie")
    d.Add "chk.audit", U("Audit rysunk\u00f3w")
    d.Add "lbl.options", U("Opcje wyj\u015bciowe:")
    d.Add "opt.single", "Pojedynczy PDF (zbiorczy)"
    d.Add "opt.multi", "Pliki indywidualne"
    d.Add "chk.open", U("Otw\u00f3rz folder po zako\u0144czeniu")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", "Jasny"
    d.Add "theme.dark", "Ciemny"
    d.Add "theme.system", "Systemowy"
    d.Add "msg.select.folder", "Wybierz folder do skanowania."
    d.Add "msg.no.idw", U("Nie znaleziono rysunk\u00f3w Inventor (.idw).")
    d.Add "msg.confirm.start", U("Czy rozpocz\u0105\u0107 proces dla {0} plik\u00f3w?")
    d.Add "msg.process.done", U("Proces zako\u0144czony pomy\u015blnie.")
    d.Add "msg.error.title", U("B\u0142\u0105d systemu")
    d.Add "log.start", U("Sesja rozpocz\u0119ta: {0}")
    d.Add "log.end", U("Sesja zako\u0144czona: {0}")
    d.Add "log.save", "Log zapisany w: {0}"
    d.Add "scan.sum.drawings", U("{0} rysunk\u00f3w")
    d.Add "scan.sum.subfolders", "z podfolderami"
    d.Add "scan.sum.ok", U("{0} aktualnych")
    d.Add "scan.sum.all.current", U("wszystkie aktualne")
    d.Add "scan.sum.no.idw", U("Nie znaleziono rysunk\u00f3w")
    d.Add "scan.sum.missing", U("{0} brakuj\u0105cych")
    d.Add "scan.sum.stale", U("{0} nieaktualnych")
    d.Add "scan.sum.mismatch", U("{0} niezgodnych")
    d.Add "scan.sum.to_create", U("{0} do utworzenia")
    d.Add "scan.sum.to_recreate", U("{0} do ponownego utworzenia")
    d.Add "scan.sum.force_suffix", U("(wymuszone)")
    d.Add "scan.validate.error", U("B\u0142\u0105d walidacji folderu.")
    d.Add "scan.sum.errors", U("{0} b\u0142\u0119d\u00f3w")
    d.Add "scan.no.folder", "Nie wybrano folderu."
    d.Add "scan.no.folder.exists", "Folder nie istnieje."
    d.Add "scan.path.ok", U("Folder znaleziony \u2014 skanowanie...")
    d.Add "scan.path.too.long", U("\u015acie\u017cka przekracza limit {0} znak\u00f3w.")
    d.Add "progress.processing", "Przetwarzanie..."
    d.Add "progress.init", "Inicjalizacja..."
    d.Add "progress.cancel", "Anuluj"
    d.Add "progress.cancelling", "Anulowanie..."
    d.Add "scan.models.detail", "Z rysunkiem: {0} | Bez: {1}"
    d.Add "status.idle", "Gotowy"
    d.Add "status.scanning", "Skanowanie folderu..."
    d.Add "status.publishing", "Publikowanie: {0} z {1}..."
    d.Add "pdf.err.open", U("Nie mo\u017cna otworzy\u0107 pliku IDW.")
    d.Add "pdf.err.pagematch", "PDF ma {0} stron, oczekiwano {1}."
    d.Add "pdf.err.noext", U("Nieprawid\u0142owa \u015bcie\u017cka (brak rozszerzenia): {0}")
    d.Add "pdf.skipped.exists", U("Pomini\u0119to (istnieje)")
    d.Add "pdf.skipped.exists.pages", U("Pomini\u0119to (istnieje) \u00b7 {0}/{1} stron")
    d.Add "pdf.err.addin", "Dodatek PDF Translator nie znaleziony."
    d.Add "pdf.err.nofile", U("Eksport PDF nie utworzy\u0142 pliku.")
    d.Add "pdf.err.export", U("Eksport nie powi\u00f3d\u0142 si\u0119: {0}")
    d.Add "pdf.err.bak.delete", U("Nie mo\u017cna usun\u0105\u0107 starej kopii zapasowej: {0}")
    d.Add "pdf.err.bak.move", U("Nie mo\u017cna wykona\u0107 kopii zapasowej PDF: {0}")
    d.Add "pdf.err.write", U("Nie mo\u017cna zapisa\u0107 PDF: {0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 Raport eksportu")
    d.Add "report.hdr.mode", "Tryb: {0}"
    d.Add "report.hdr.folder", "Folder: {0}"
    d.Add "report.hdr.subfolders", "Podfoldery"
    d.Add "report.hdr.force", "Wymuszone nadpisanie"
    d.Add "report.hdr.generated", "Wygenerowano: {0}"
    d.Add "report.state.on", U("W\u0142.")
    d.Add "report.state.off", U("Wy\u0142.")
    d.Add "report.hdr.pdf.section", "EKSPORT PDF"
    d.Add "report.hdr.pdf.detail", U("Szczeg\u00f3\u0142y:")
    d.Add "report.created", "Utworzono"
    d.Add "report.overwritten", "Nadpisano"
    d.Add "report.skipped", U("Pomini\u0119to")
    d.Add "report.errors", U("B\u0142\u0119dy")
    d.Add "report.time", "Czas"
    d.Add "report.empty", "pusty"
    d.Add "report.pages", "stron"
    d.Add "report.no.views", U("Brak widok\u00f3w")
    d.Add "report.tag.created", "[Utworzono]"
    d.Add "report.tag.overwritten", "[Nadpisano]"
    d.Add "report.tag.skipped", U("[Pomini\u0119to]")
    d.Add "report.tag.error", U("[B\u0142\u0105d]")
    d.Add "report.tag.hasdrawing", "[Ma rysunek]"
    d.Add "report.tag.nodrawing", "[Brak rysunku]"
    d.Add "report.hdr.audit.section", U("AUDYT RYSUNK\u00d3W")
    d.Add "report.hdr.audit.detail", U("Szczeg\u00f3\u0142y:")
    d.Add "report.audit.title", U("Audyt rysunk\u00f3w")
    d.Add "errdlg.retry.prompt", U("Pon\u00f3w = pon\u00f3w pr\u00f3b\u0119 | Ignoruj = pomi\u0144 | Przerwij = zatrzymaj")
    d.Add "tip.run", U("Rozpocznij eksport PDF w trybie wsadowym")
    d.Add "tip.theme", U("Prze\u0142\u0105czaj motywy (Inventor / Systemowy / Jasny / Ciemny)")
    d.Add "tip.reset", U("Przywr\u00f3\u0107 wszystkie opcje domy\u015blne")
    d.Add "tip.cancel", U("Anuluj bie\u017c\u0105c\u0105 operacj\u0119 (ESC)")
    d.Add "tip.filter.nodrawing", U("Poka\u017c tylko modele bez rysunku")
    d.Add "tip.list", U("Kliknij dwukrotnie, aby otworzy\u0107 PDF; kliknij, aby zobaczy\u0107 szczeg\u00f3\u0142y")
    d.Add "tip.subfolders", U("Rekurencyjnie skanuj podfoldery (pomija OldVersions)")
    d.Add "tip.force", U("Utw\u00f3rz PDF ponownie, nawet je\u015bli jest nowszy ni\u017c IDW")
    d.Add "tip.audit", U("Sprawd\u017a modele (.ipt/.iam) pod k\u0105tem brakuj\u0105cych rysunk\u00f3w")
    d.Add "tip.retry", U("Pon\u00f3w wszystkie pliki z b\u0142\u0119dem")
    d.Add "tip.export", U("Zapisz raport do pliku i otw\u00f3rz w Notatniku")
    d.Add "report.msg.close_blocked", U("Nie mo\u017cna zamkn\u0105\u0107 podczas przetwarzania.")
    d.Add "report.msg.open_failed", U("Nie mo\u017cna otworzy\u0107: plik nie znaleziony lub akcja zablokowana.")
    d.Add "report.msg.select_row", U("Wybierz wiersz, aby zobaczy\u0107 szczeg\u00f3\u0142y  \u00b7  Kliknij dwukrotnie, aby otworzy\u0107 PDF")
    d.Add "report.msg.retry.none", U("Pon\u00f3w: brak plik\u00f3w w stanie b\u0142\u0119du.")
    d.Add "report.msg.retry.error", U("B\u0142\u0105d ponowienia: {0}")
    d.Add "report.msg.retry.partial", U("Pon\u00f3w: {0} udanych, {1} nadal w b\u0142\u0119dzie.")
    d.Add "report.msg.retry.all_fixed", U("Pon\u00f3w: {0} udanych, wszystkie naprawione.")
    d.Add "report.msg.copied", U("Skopiowano!")
    d.Add "report.msg.exported", U("Raport otwarty w Notatniku.")
    d.Add "report.msg.export_failed", U("Nie uda\u0142o si\u0119 zapisa\u0107 raportu.")
    d.Add "reset.msg.all_defaults", U("Wszystkie ustawienia przywrocone.")
    d.Add "reset.lbl.registry", "Rejestr"
    d.Add "reset.lbl.sandbox", "Sandbox"
    d.Add "reset.lbl.state", "Stan"
    d.Add "reset.state.deleted", U("Usunieto")
    d.Add "reset.state.reinitialized", U("Zainicjalizowano ponownie")
    d.Add "reset.msg.failed", U("Reset nie powiodl sie")
    d.Add "reset.msg.registry_complete", U("Rejestr zresetowany.")
    d.Add "reset.msg.sandbox_complete", U("Sandbox zresetowany")
    d.Add "scan.badge.tooltip", U("Razem: {0} | Aktualne: {1} | Brakuj\u0105ce: {2} | Nieaktualne: {3} | Niezgodne: {4} | B\u0142\u0119dy: {5}")
    d.Add "report.msg.opening", U("Otwieranie pliku IDW...")
    d.Add "report.msg.updating", U("Aktualizacja widok\u00f3w...")
    d.Add "report.msg.exporting", U("Eksport PDF...")
    d.Add "msg.no.write.access", U("Nie mo\u017cna zapisa\u0107 w folderze sandbox. Sprawd\u017a uprawnienia:")
    d.Add "msg.pdf.temp_missing", U("Tymczasowy PDF znikn\u0105\u0142 podczas eksportu (mo\u017cliwa kwarantanna antywirusowa lub blokada pliku): {0}")
    d.Add "msg.batch.abort.network", U("Partia przerwana: {0} kolejnych b\u0142\u0119d\u00f3w sieci. Sprawd\u017a po\u0142\u0105czenie i spr\u00f3buj ponownie.")
    d.Add "msg.already.running", U("Batch PDF Publisher jest ju\u017c uruchomiony. Zamknij drug\u0105 sesj\u0119 i spr\u00f3buj ponownie.")
    d.Add "report.search.placeholder", U("Szukaj \u015bcie\u017cki...")
    d.Add "report.search.label", "Szukaj:"
    d.Add "scan.collecting", U("Zbieranie plik\u00f3w...")
    d.Add "scan.analysing", "Analiza: {0} / {1}..."
    d.Add "scan.auditing.start", U("Sprawdzanie modeli...")
    d.Add "scan.auditing", "Audyt: {0} / {1}..."
    d.Add "scan.cancelled", U("Anulowano")
    d.Add "tip.search", "Filtruj wiersze raportu wed\u0142ug \u015bcie\u017cki"
    d.Add "msg.confirm.reset", U("Przywr\u00f3ci\u0107 wszystkie opcje domy\u015blne?")
    d.Add "scan.disk.insufficient", U("Za ma\u0142o miejsca na dysku. Wolne: {0}")
    d.Add "scan.title.no_folder", "Wybierz folder"
    d.Add "summary.hdr.models_full", "{0} modeli ({1} z\u0142o\u017ce\u0144, {2} cz\u0119\u015bci)"
    d.Add "summary.hdr.with_idw", "{0} z rysunkiem"
    d.Add "summary.hdr.no_idw", "{0} bez rysunku"
    d.Add "errdlg.label.context", "Kontekst"
    d.Add "errdlg.label.time", "Czas"
    d.Add "errdlg.label.log", "Dziennik"
    d.Add "errdlg.label.stack", "Stos"
    d.Add "pdf.err.prepare", U("Przygotowanie nie powiod\u0142o si\u0119: {0}")
    d.Add "pdf.err.header", U("Wyeksportowany tymczasowy PDF nie przeszed\u0142 walidacji nag\u0142\u00f3wka")
    d.Add "pdf.err.unknown", U("Nieznany b\u0142\u0105d")
    d.Add "report.filter.all", "Wszystkie"
    d.Add "report.filter.err", U("B\u0142\u0119dy")
    d.Add "report.filter.skip", U("Pomini\u0119te")
    d.Add "report.filter.ok", "OK"
    d.Add "report.filter.noidw", "Bez .idw"
    d.Add "scan.reset.sure", U("Pewny?")
    d.Add "scan.reset.done", "Gotowe"
    d.Add "report.form.title", "Raport przetwarzania"
    d.Add "report.hdr.sub", "Podfoldery"
    d.Add "report.hdr.force_short", U("Wymu\u015b")
    d.Add "report.hdr.audit", "Audyt"
    d.Add "report.summary.full", U("{0} utworzonych \u00b7 {1} nadpisanych \u00b7 {2} pomini\u0119tych \u00b7 {3} b\u0142\u0119d\u00f3w \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} pustych)")
    d.Add "scan.browse_title", U("Wybierz folder \u017ar\u00f3d\u0142owy")
End Sub

Private Sub L_FillCS(d As Object)
    d.Add "title", U("Hromadn\u00e1 publikace do PDF")
    d.Add "btn.run", "SPUSTIT"
    d.Add "btn.cancel", "Storno"
    d.Add "btn.browse", U("Proch\u00e1zet...")
    d.Add "btn.copy", U("Kop\u00edrovat")
    d.Add "btn.help", U("N\u00e1pov\u011bda")
    d.Add "btn.settings", "Nastaven\u00ed"
    d.Add "btn.back", U("Zp\u011bt")
    d.Add "btn.export", "Exportovat"
    d.Add "btn.retry", "Opakovat"
    d.Add "btn.reset", "Resetovat"
    d.Add "tab.scan", "Skenovat"
    d.Add "tab.settings", "Nastaven\u00ed"
    d.Add "tab.log", "Protokol"
    d.Add "settings.lang", "Jazyk:"
    d.Add "settings.lang.auto", "Auto (host)"
    d.Add "settings.loglevel", U("\u00darove\u0148 logu:")
    d.Add "settings.loglevel.debug", "Debug"
    d.Add "settings.loglevel.info", "Info"
    d.Add "settings.loglevel.warn", U("Varov\u00e1n\u00ed")
    d.Add "settings.loglevel.error", "Pouze chyby"
    d.Add "lbl.folder", U("Zdrojov\u00e1 slo\u017eka:")
    d.Add "lbl.subfolders", U("V\u010detn\u011b podslo\u017eek")
    d.Add "chk.force", U("Vynutit p\u0159egenerov\u00e1n\u00ed")
    d.Add "chk.audit", U("Audit v\u00fdkres\u016f")
    d.Add "lbl.options", U("Mo\u017enosti v\u00fdstupu:")
    d.Add "opt.single", U("Jedin\u00e9 PDF (slou\u010den\u00e9)")
    d.Add "opt.multi", U("Jednotliv\u00e9 soubory")
    d.Add "chk.open", U("Otev\u0159\u00edt slo\u017eku po dokon\u010den\u00ed")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", U("Sv\u011btl\u00fd")
    d.Add "theme.dark", U("Tmav\u00fd")
    d.Add "theme.system", U("Syst\u00e9mov\u00fd")
    d.Add "msg.select.folder", U("Vyberte slo\u017eku ke skenov\u00e1n\u00ed.")
    d.Add "msg.no.idw", U("Nebyly nalezeny \u017e\u00e1dn\u00e9 v\u00fdkresy Inventor (.idw).")
    d.Add "msg.confirm.start", U("Chcete spustit proces pro {0} soubor\u016f?")
    d.Add "msg.process.done", U("Proces byl \u00fasp\u011b\u0161n\u011b dokon\u010den.")
    d.Add "msg.error.title", U("Syst\u00e9mov\u00e1 chyba")
    d.Add "log.start", U("Relace zah\u00e1jena: {0}")
    d.Add "log.end", U("Relace ukon\u010dena: {0}")
    d.Add "log.save", U("Protokol ulo\u017een do: {0}")
    d.Add "scan.sum.drawings", U("{0} v\u00fdkres\u016f")
    d.Add "scan.sum.subfolders", U("v\u010d. podslo\u017eek")
    d.Add "scan.sum.ok", U("{0} aktu\u00e1ln\u00edch")
    d.Add "scan.sum.all.current", U("v\u0161echny aktu\u00e1ln\u00ed")
    d.Add "scan.sum.no.idw", U("Nenalezeny \u017e\u00e1dn\u00e9 v\u00fdkresy")
    d.Add "scan.sum.missing", U("{0} chyb\u011bj\u00edc\u00edch")
    d.Add "scan.sum.stale", U("{0} zastaral\u00fdch")
    d.Add "scan.sum.mismatch", U("{0} nesouladn\u00fdch")
    d.Add "scan.sum.to_create", U("{0} k vytvo\u0159en\u00ed")
    d.Add "scan.sum.to_recreate", U("{0} k op\u011btovn\u00e9mu vytvo\u0159en\u00ed")
    d.Add "scan.sum.force_suffix", U("(vynuceno)")
    d.Add "scan.validate.error", U("Chyba p\u0159i ov\u011b\u0159ov\u00e1n\u00ed slo\u017eky.")
    d.Add "scan.sum.errors", U("{0} chyb")
    d.Add "scan.no.folder", U("Nebyla vybr\u00e1na \u017e\u00e1dn\u00e1 slo\u017eka.")
    d.Add "scan.no.folder.exists", U("Slo\u017eka neexistuje.")
    d.Add "scan.path.ok", U("Slo\u017eka nalezena \u2014 prob\u00edh\u00e1 sken...")
    d.Add "scan.path.too.long", U("Cesta p\u0159ekra\u010duje limit {0} znak\u016f.")
    d.Add "progress.processing", U("Zpracov\u00e1n\u00ed...")
    d.Add "progress.init", "Inicializace..."
    d.Add "progress.cancel", "Storno"
    d.Add "progress.cancelling", U("Ru\u0161en\u00ed...")
    d.Add "scan.models.detail", U("S v\u00fdkresem: {0} | Bez: {1}")
    d.Add "status.idle", U("P\u0159ipraveno")
    d.Add "status.scanning", U("Skenov\u00e1n\u00ed slo\u017eky...")
    d.Add "status.publishing", U("Publikov\u00e1n\u00ed: {0} z {1}...")
    d.Add "pdf.err.open", U("Nelze otev\u0159\u00edt soubor IDW.")
    d.Add "pdf.err.pagematch", U("PDF m\u00e1 {0} str\u00e1nek, o\u010dek\u00e1v\u00e1no {1}.")
    d.Add "pdf.err.noext", U("Neplatn\u00e1 cesta (bez p\u0159\u00edpony): {0}")
    d.Add "pdf.skipped.exists", U("P\u0159esko\u010deno (existuje)")
    d.Add "pdf.skipped.exists.pages", U("P\u0159esko\u010deno (existuje) \u00b7 {0}/{1} str\u00e1nek")
    d.Add "pdf.err.addin", U("Dopl\u0148\u011bk PDF Translator nenalezen.")
    d.Add "pdf.err.nofile", U("Export PDF nevytvo\u0159il \u017e\u00e1dn\u00fd soubor.")
    d.Add "pdf.err.export", U("Export selhal: {0}")
    d.Add "pdf.err.bak.delete", U("Nelze odstranit starou z\u00e1lohu: {0}")
    d.Add "pdf.err.bak.move", U("Nelze z\u00e1lohovat existuj\u00edc\u00ed PDF: {0}")
    d.Add "pdf.err.write", U("Nelze zapsat PDF: {0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 Exportn\u00ed zpr\u00e1va")
    d.Add "report.hdr.mode", U("Re\u017eim: {0}")
    d.Add "report.hdr.folder", U("Slo\u017eka: {0}")
    d.Add "report.hdr.subfolders", U("Podslo\u017eky")
    d.Add "report.hdr.force", U("Vynucen\u00e9 p\u0159episov\u00e1n\u00ed")
    d.Add "report.hdr.generated", U("Vygenerov\u00e1no: {0}")
    d.Add "report.state.on", "Ano"
    d.Add "report.state.off", "Ne"
    d.Add "report.hdr.pdf.section", "EXPORT PDF"
    d.Add "report.hdr.pdf.detail", "Detaily:"
    d.Add "report.created", U("Vytvo\u0159eno")
    d.Add "report.overwritten", U("P\u0159eps\u00e1no")
    d.Add "report.skipped", U("P\u0159esko\u010deno")
    d.Add "report.errors", "Chyby"
    d.Add "report.time", U("\u010cas")
    d.Add "report.empty", U("pr\u00e1zdn\u00e9")
    d.Add "report.pages", U("str\u00e1nek")
    d.Add "report.no.views", U("\u017d\u00e1dn\u00e9 pohledy")
    d.Add "report.tag.created", U("[Vytvo\u0159eno]")
    d.Add "report.tag.overwritten", U("[P\u0159eps\u00e1no]")
    d.Add "report.tag.skipped", U("[P\u0159esko\u010deno]")
    d.Add "report.tag.error", "[Chyba]"
    d.Add "report.tag.hasdrawing", U("[M\u00e1 v\u00fdkres]")
    d.Add "report.tag.nodrawing", U("[Bez v\u00fdkresu]")
    d.Add "report.hdr.audit.section", U("AUDIT V\u00ddKRES\u016e")
    d.Add "report.hdr.audit.detail", "Detaily:"
    d.Add "report.audit.title", U("Audit v\u00fdkres\u016f")
    d.Add "errdlg.retry.prompt", U("Zkusit znovu = opakovat | Ignorovat = p\u0159esko\u010dit | P\u0159eru\u0161it = zastavit")
    d.Add "tip.run", U("Spustit d\u00e1vkov\u00fd export PDF")
    d.Add "tip.theme", U("Proch\u00e1zet motivy (Inventor / Syst\u00e9m / Sv\u011btl\u00fd / Tmav\u00fd)")
    d.Add "tip.reset", U("Obnovit v\u0161echny mo\u017enosti na v\u00fdchoz\u00ed hodnoty")
    d.Add "tip.cancel", U("Zru\u0161it aktu\u00e1ln\u00ed operaci (ESC)")
    d.Add "tip.filter.nodrawing", U("Zobrazit pouze modely bez v\u00fdkresu")
    d.Add "tip.list", U("Dvojklikem otev\u0159ete PDF; kliknut\u00edm zobraz\u00edte podrobnosti")
    d.Add "tip.subfolders", U("Rekurzivn\u011b prohledat podslo\u017eky (p\u0159esko\u010d\u00ed OldVersions)")
    d.Add "tip.force", U("Znovu vytvo\u0159it PDF i kdy\u017e je nov\u011bj\u0161\u00ed ne\u017e IDW")
    d.Add "tip.audit", U("Zkontrolovat modely (.ipt/.iam) na chyb\u011bj\u00edc\u00ed v\u00fdkresy")
    d.Add "tip.retry", U("Zopakovat v\u0161echny soubory s chybou")
    d.Add "tip.export", U("Ulo\u017eit zpr\u00e1vu do souboru a otev\u0159\u00edt v Pozn\u00e1mkov\u00e9m bloku")
    d.Add "report.msg.close_blocked", U("Nelze zav\u0159\u00edt b\u011bhem zpracov\u00e1n\u00ed.")
    d.Add "report.msg.open_failed", U("Nelze otev\u0159\u00edt: soubor nenalezen nebo akce blokov\u00e1na.")
    d.Add "report.msg.select_row", U("Vyberte \u0159\u00e1dek pro podrobnosti  \u00b7  Dvojklik pro otev\u0159en\u00ed PDF")
    d.Add "report.msg.retry.none", U("Zkusit znovu: \u017e\u00e1dn\u00e9 soubory ve stavu chyby.")
    d.Add "report.msg.retry.error", U("Chyba opakov\u00e1n\u00ed: {0}")
    d.Add "report.msg.retry.partial", U("Opakov\u00e1n\u00ed: {0} \u00fasp\u011b\u0161n\u00fdch, {1} st\u00e1le s chybou.")
    d.Add "report.msg.retry.all_fixed", U("Opakov\u00e1n\u00ed: {0} \u00fasp\u011b\u0161n\u00fdch, v\u0161e opraveno.")
    d.Add "report.msg.copied", U("Zkop\u00edrov\u00e1no!")
    d.Add "report.msg.exported", U("Zpr\u00e1va otev\u0159ena v Pozn\u00e1mkov\u00e9m bloku.")
    d.Add "report.msg.export_failed", U("Nepoda\u0159ilo se ulo\u017eit zpr\u00e1vu.")
    d.Add "reset.msg.all_defaults", U("Vsechna nastaveni obnovena.")
    d.Add "reset.lbl.registry", "Registr"
    d.Add "reset.lbl.sandbox", "Sandbox"
    d.Add "reset.lbl.state", "Stav"
    d.Add "reset.state.deleted", U("Smazano")
    d.Add "reset.state.reinitialized", U("Reinicializovano")
    d.Add "reset.msg.failed", U("Reset selhal")
    d.Add "reset.msg.registry_complete", U("Registr resetovan.")
    d.Add "reset.msg.sandbox_complete", U("Sandbox resetovan")
    d.Add "scan.badge.tooltip", U("Celkem: {0} | Aktu\u00e1ln\u00ed: {1} | Chyb\u011bj\u00edc\u00ed: {2} | Zastaral\u00e9: {3} | Nesoulad: {4} | Chyby: {5}")
    d.Add "report.msg.opening", U("Otev\u00edr\u00e1n\u00ed souboru IDW...")
    d.Add "report.msg.updating", U("Aktualizace pohled\u016f...")
    d.Add "report.msg.exporting", "Export PDF..."
    d.Add "msg.no.write.access", U("Nelze zapisovat do slo\u017eky sandbox. Zkontrolujte opr\u00e1vn\u011bn\u00ed:")
    d.Add "msg.pdf.temp_missing", U("Do\u010dasn\u00fd PDF b\u011bhem exportu zmizel (mo\u017en\u00e1 karant\u00e9na antiviru nebo uzam\u010den\u00ed souboru): {0}")
    d.Add "msg.batch.abort.network", U("D\u00e1vka p\u0159eru\u0161ena: {0} n\u00e1sledn\u00fdch chyb s\u00edt\u011b. Zkontrolujte p\u0159ipojen\u00ed a zkuste to znovu.")
    d.Add "msg.already.running", U("Batch PDF Publisher ji\u017e b\u011b\u017e\u00ed. Zav\u0159ete druhou relaci a zkuste to znovu.")
    d.Add "report.search.placeholder", U("Hledat cestu...")
    d.Add "report.search.label", "Hledat:"
    d.Add "scan.collecting", U("Sb\u00edr\u00e1n\u00ed soubor\u016f...")
    d.Add "scan.analysing", U("Anal\u00fdza: {0} / {1}...")
    d.Add "scan.auditing.start", U("Kontrola model\u016f...")
    d.Add "scan.auditing", "Audit: {0} / {1}..."
    d.Add "scan.cancelled", U("Zru\u0161eno")
    d.Add "tip.search", U("Filtrovat \u0159\u00e1dky sestavy podle cesty")
    d.Add "msg.confirm.reset", U("Obnovit v\u0161echny mo\u017enosti na v\u00fdchoz\u00ed hodnoty?")
    d.Add "scan.disk.insufficient", U("Nedostatek m\u00edsta na disku. Voln\u00e9: {0}")
    d.Add "scan.title.no_folder", "Vyberte slo\u017eku"
    d.Add "summary.hdr.models_full", "{0} model\u016f ({1} sestav, {2} sou\u010d\u00e1st\u00ed)"
    d.Add "summary.hdr.with_idw", "{0} s v\u00fdkresem"
    d.Add "summary.hdr.no_idw", "{0} bez v\u00fdkresu"
    d.Add "errdlg.label.context", "Kontext"
    d.Add "errdlg.label.time", U("\u010cas")
    d.Add "errdlg.label.log", "Protokol"
    d.Add "errdlg.label.stack", "Z\u00e1sobn\u00edk"
    d.Add "pdf.err.prepare", U("P\u0159\u00edprava selhala: {0}")
    d.Add "pdf.err.header", U("Exportovan\u00fd do\u010dasn\u00fd PDF nepro\u0161el validac\u00ed z\u00e1hlav\u00ed")
    d.Add "pdf.err.unknown", U("Nezn\u00e1m\u00e1 chyba")
    d.Add "report.filter.all", U("V\u0161e")
    d.Add "report.filter.err", "Chyby"
    d.Add "report.filter.skip", U("P\u0159esko\u010deno")
    d.Add "report.filter.ok", "OK"
    d.Add "report.filter.noidw", "Bez .idw"
    d.Add "scan.reset.sure", U("Jist\u011b?")
    d.Add "scan.reset.done", "Hotovo"
    d.Add "report.form.title", "Protokol zpracov\u00e1n\u00ed"
    d.Add "report.hdr.sub", U("Podslo\u017eky")
    d.Add "report.hdr.force_short", "Vynutit"
    d.Add "report.hdr.audit", "Audit"
    d.Add "report.summary.full", U("{0} vytvo\u0159eno \u00b7 {1} p\u0159eps\u00e1no \u00b7 {2} p\u0159esko\u010deno \u00b7 {3} chyb \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} pr\u00e1zdn\u00fdch)")
    d.Add "scan.browse_title", U("Vybrat zdrojovou slo\u017eku")
End Sub

Private Sub L_FillJA(d As Object)
    d.Add "title", U("PDF \u4e00\u62ec\u767a\u884c")
    d.Add "btn.run", U("\u5b9f\u884c")
    d.Add "btn.cancel", U("\u30ad\u30e3\u30f3\u30bb\u30eb")
    d.Add "btn.browse", U("\u53c2\u7167...")
    d.Add "btn.copy", U("\u30b3\u30d4\u30fc")
    d.Add "btn.help", U("\u30d8\u30eb\u30d7")
    d.Add "btn.settings", U("\u8a2d\u5b9a")
    d.Add "btn.back", U("\u623b\u308b")
    d.Add "btn.export", U("\u30a8\u30af\u30b9\u30dd\u30fc\u30c8")
    d.Add "btn.retry", U("\u518d\u8a66\u884c")
    d.Add "btn.reset", U("\u30ea\u30bb\u30c3\u30c8")
    d.Add "tab.scan", U("\u30b9\u30ad\u30e3\u30f3")
    d.Add "tab.settings", U("\u8a2d\u5b9a")
    d.Add "tab.log", U("\u30ed\u30b0")
    d.Add "settings.lang", U("\u8a00\u8a9e:")
    d.Add "settings.lang.auto", U("\u81ea\u52d5 (\u30db\u30b9\u30c8)")
    d.Add "settings.loglevel", U("\u30ed\u30b0\u30ec\u30d9\u30eb:")
    d.Add "settings.loglevel.debug", U("\u30c7\u30d0\u30c3\u30b0")
    d.Add "settings.loglevel.info", U("\u60c5\u5831")
    d.Add "settings.loglevel.warn", U("\u8b66\u544a")
    d.Add "settings.loglevel.error", U("\u30a8\u30e9\u30fc\u306e\u307f")
    d.Add "lbl.folder", U("\u30bd\u30fc\u30b9\u30d5\u30a9\u30eb\u30c0\u30fc:")
    d.Add "lbl.subfolders", U("\u30b5\u30d6\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u542b\u3081\u308b")
    d.Add "chk.force", U("\u518d\u751f\u6210\u3092\u5f37\u5236")
    d.Add "chk.audit", U("\u56f3\u9762\u76e3\u67fb")
    d.Add "lbl.options", U("\u51fa\u529b\u30aa\u30d7\u30b7\u30e7\u30f3:")
    d.Add "opt.single", U("\u5358\u4e00 PDF (\u4e00\u62ec)")
    d.Add "opt.multi", U("\u500b\u5225\u30d5\u30a1\u30a4\u30eb")
    d.Add "chk.open", U("\u5b8c\u4e86\u5f8c\u306b\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u958b\u304f")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", U("\u30e9\u30a4\u30c8")
    d.Add "theme.dark", U("\u30c0\u30fc\u30af")
    d.Add "theme.system", U("\u30b7\u30b9\u30c6\u30e0")
    d.Add "msg.select.folder", U("\u30b9\u30ad\u30e3\u30f3\u3059\u308b\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002")
    d.Add "msg.no.idw", U("Inventor \u56f3\u9762 (.idw) \u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002")
    d.Add "msg.confirm.start", U("{0} \u30d5\u30a1\u30a4\u30eb\u306e\u51e6\u7406\u3092\u958b\u59cb\u3057\u307e\u3059\u304b\uff1f")
    d.Add "msg.process.done", U("\u51e6\u7406\u304c\u6b63\u5e38\u306b\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002")
    d.Add "msg.error.title", U("\u30b7\u30b9\u30c6\u30e0\u30a8\u30e9\u30fc")
    d.Add "log.start", U("\u30bb\u30c3\u30b7\u30e7\u30f3\u958b\u59cb: {0}")
    d.Add "log.end", U("\u30bb\u30c3\u30b7\u30e7\u30f3\u7d42\u4e86: {0}")
    d.Add "log.save", U("\u30ed\u30b0\u3092\u4fdd\u5b58\u3057\u307e\u3057\u305f: {0}")
    d.Add "scan.sum.drawings", U("{0} \u56f3\u9762")
    d.Add "scan.sum.subfolders", U("\u30b5\u30d6\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u542b\u3080")
    d.Add "scan.sum.ok", U("{0} \u6700\u65b0")
    d.Add "scan.sum.all.current", U("\u3059\u3079\u3066\u6700\u65b0")
    d.Add "scan.sum.no.idw", U("\u56f3\u9762\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093")
    d.Add "scan.sum.missing", U("{0} \u6b20\u843d")
    d.Add "scan.sum.stale", U("{0} \u53e4\u3044")
    d.Add "scan.sum.mismatch", U("{0} \u4e0d\u4e00\u81f4")
    d.Add "scan.sum.to_create", U("{0} \u4f5c\u6210")
    d.Add "scan.sum.to_recreate", U("{0} \u518d\u4f5c\u6210")
    d.Add "scan.sum.force_suffix", U("(\u5f37\u5236)")
    d.Add "scan.validate.error", U("\u30d5\u30a9\u30eb\u30c0\u30fc\u691c\u8a3c\u30a8\u30e9\u30fc\u3002")
    d.Add "scan.sum.errors", U("{0} \u30a8\u30e9\u30fc")
    d.Add "scan.no.folder", U("\u30d5\u30a9\u30eb\u30c0\u30fc\u304c\u9078\u629e\u3055\u308c\u3066\u3044\u307e\u305b\u3093\u3002")
    d.Add "scan.no.folder.exists", U("\u30d5\u30a9\u30eb\u30c0\u30fc\u304c\u5b58\u5728\u3057\u307e\u305b\u3093\u3002")
    d.Add "scan.path.ok", U("\u30d5\u30a9\u30eb\u30c0\u30fc\u304c\u898b\u3064\u304b\u308a\u307e\u3057\u305f \u2014 \u30b9\u30ad\u30e3\u30f3\u4e2d...")
    d.Add "scan.path.too.long", U("\u30d1\u30b9\u304c {0} \u6587\u5b57\u306e\u5236\u9650\u3092\u8d85\u3048\u3066\u3044\u307e\u3059\u3002")
    d.Add "progress.processing", U("\u51e6\u7406\u4e2d...")
    d.Add "progress.init", U("\u521d\u671f\u5316\u4e2d...")
    d.Add "progress.cancel", U("\u30ad\u30e3\u30f3\u30bb\u30eb")
    d.Add "progress.cancelling", U("\u30ad\u30e3\u30f3\u30bb\u30eb\u4e2d...")
    d.Add "scan.models.detail", U("\u56f3\u9762\u3042\u308a: {0} | \u306a\u3057: {1}")
    d.Add "status.idle", U("\u6e96\u5099\u5b8c\u4e86")
    d.Add "status.scanning", U("\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u30b9\u30ad\u30e3\u30f3\u4e2d...")
    d.Add "status.publishing", U("\u767a\u884c\u4e2d: {0} / {1}...")
    d.Add "pdf.err.open", U("IDW \u30d5\u30a1\u30a4\u30eb\u3092\u958b\u3051\u307e\u305b\u3093\u3002")
    d.Add "pdf.err.pagematch", U("PDF \u306f {0} \u30da\u30fc\u30b8\u3067\u3059\u304c\u3001{1} \u30da\u30fc\u30b8\u3092\u671f\u5f85\u3057\u307e\u3057\u305f\u3002")
    d.Add "pdf.err.noext", U("\u7121\u52b9\u306a\u30d1\u30b9 (\u62e1\u5f35\u5b50\u306a\u3057): {0}")
    d.Add "pdf.skipped.exists", U("\u30b9\u30ad\u30c3\u30d7 (\u5b58\u5728)")
    d.Add "pdf.skipped.exists.pages", U("\u30b9\u30ad\u30c3\u30d7 (\u5b58\u5728) \u00b7 {0}/{1} \u30da\u30fc\u30b8")
    d.Add "pdf.err.addin", U("PDF Translator \u30a2\u30c9\u30a4\u30f3\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002")
    d.Add "pdf.err.nofile", U("PDF \u30a8\u30af\u30b9\u30dd\u30fc\u30c8\u3067\u30d5\u30a1\u30a4\u30eb\u304c\u4f5c\u6210\u3055\u308c\u307e\u305b\u3093\u3067\u3057\u305f\u3002")
    d.Add "pdf.err.export", U("\u30a8\u30af\u30b9\u30dd\u30fc\u30c8\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {0}")
    d.Add "pdf.err.bak.delete", U("\u53e4\u3044\u30d0\u30c3\u30af\u30a2\u30c3\u30d7\u3092\u524a\u9664\u3067\u304d\u307e\u305b\u3093: {0}")
    d.Add "pdf.err.bak.move", U("\u65e2\u5b58\u306e PDF \u3092\u30d0\u30c3\u30af\u30a2\u30c3\u30d7\u3067\u304d\u307e\u305b\u3093: {0}")
    d.Add "pdf.err.write", U("PDF \u3092\u66f8\u304d\u8fbc\u3081\u307e\u305b\u3093: {0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 \u30a8\u30af\u30b9\u30dd\u30fc\u30c8\u30ec\u30dd\u30fc\u30c8")
    d.Add "report.hdr.mode", U("\u30e2\u30fc\u30c9: {0}")
    d.Add "report.hdr.folder", U("\u30d5\u30a9\u30eb\u30c0\u30fc: {0}")
    d.Add "report.hdr.subfolders", U("\u30b5\u30d6\u30d5\u30a9\u30eb\u30c0\u30fc")
    d.Add "report.hdr.force", U("\u5f37\u5236\u4e0a\u66f8\u304d")
    d.Add "report.hdr.generated", U("\u751f\u6210: {0}")
    d.Add "report.state.on", U("\u30aa\u30f3")
    d.Add "report.state.off", U("\u30aa\u30d5")
    d.Add "report.hdr.pdf.section", U("PDF \u30a8\u30af\u30b9\u30dd\u30fc\u30c8")
    d.Add "report.hdr.pdf.detail", U("\u8a73\u7d30:")
    d.Add "report.created", U("\u4f5c\u6210")
    d.Add "report.overwritten", U("\u4e0a\u66f8\u304d")
    d.Add "report.skipped", U("\u30b9\u30ad\u30c3\u30d7")
    d.Add "report.errors", U("\u30a8\u30e9\u30fc")
    d.Add "report.time", U("\u6642\u9593")
    d.Add "report.empty", U("\u7a7a")
    d.Add "report.pages", U("\u30da\u30fc\u30b8")
    d.Add "report.no.views", U("\u30d3\u30e5\u30fc\u306a\u3057")
    d.Add "report.tag.created", U("[\u4f5c\u6210]")
    d.Add "report.tag.overwritten", U("[\u4e0a\u66f8\u304d]")
    d.Add "report.tag.skipped", U("[\u30b9\u30ad\u30c3\u30d7]")
    d.Add "report.tag.error", U("[\u30a8\u30e9\u30fc]")
    d.Add "report.tag.hasdrawing", U("[\u56f3\u9762\u3042\u308a]")
    d.Add "report.tag.nodrawing", U("[\u56f3\u9762\u306a\u3057]")
    d.Add "report.hdr.audit.section", U("\u56f3\u9762\u76e3\u67fb")
    d.Add "report.hdr.audit.detail", U("\u8a73\u7d30:")
    d.Add "report.audit.title", U("\u56f3\u9762\u76e3\u67fb")
    d.Add "errdlg.retry.prompt", U("\u518d\u8a66\u884c = \u518d\u8a66\u884c | \u7121\u8996 = \u30b9\u30ad\u30c3\u30d7 | \u4e2d\u6b62 = \u505c\u6b62")
    d.Add "tip.run", U("\u30d0\u30c3\u30c1 PDF \u30a8\u30af\u30b9\u30dd\u30fc\u30c8\u3092\u958b\u59cb")
    d.Add "tip.theme", U("\u30c6\u30fc\u30de\u3092\u5207\u308a\u66ff\u3048 (Inventor / \u30b7\u30b9\u30c6\u30e0 / \u30e9\u30a4\u30c8 / \u30c0\u30fc\u30af)")
    d.Add "tip.reset", U("\u3059\u3079\u3066\u306e\u30aa\u30d7\u30b7\u30e7\u30f3\u3092\u65e2\u5b9a\u306b\u623b\u3059")
    d.Add "tip.cancel", U("\u73fe\u5728\u306e\u64cd\u4f5c\u3092\u30ad\u30e3\u30f3\u30bb\u30eb (ESC)")
    d.Add "tip.filter.nodrawing", U("\u56f3\u9762\u306e\u306a\u3044\u30e2\u30c7\u30eb\u306e\u307f\u8868\u793a")
    d.Add "tip.list", U("\u30c0\u30d6\u30eb\u30af\u30ea\u30c3\u30af\u3067 PDF \u3092\u958b\u304f; \u30af\u30ea\u30c3\u30af\u3067\u8a73\u7d30\u3092\u8868\u793a")
    d.Add "tip.subfolders", U("\u30b5\u30d6\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u518d\u5e30\u7684\u306b\u30b9\u30ad\u30e3\u30f3 (OldVersions \u3092\u30b9\u30ad\u30c3\u30d7)")
    d.Add "tip.force", U("IDW \u3088\u308a\u65b0\u3057\u304f\u3066\u3082 PDF \u3092\u518d\u4f5c\u6210")
    d.Add "tip.audit", U("\u30e2\u30c7\u30eb (.ipt/.iam) \u306e\u6b20\u843d\u3057\u305f\u56f3\u9762\u3092\u78ba\u8a8d")
    d.Add "tip.retry", U("\u30a8\u30e9\u30fc\u306e\u3042\u308b\u3059\u3079\u3066\u306e\u30d5\u30a1\u30a4\u30eb\u3092\u518d\u8a66\u884c")
    d.Add "tip.export", U("\u30ec\u30dd\u30fc\u30c8\u3092\u30d5\u30a1\u30a4\u30eb\u306b\u4fdd\u5b58\u3057\u3066\u30e1\u30e2\u5e33\u3067\u958b\u304f")
    d.Add "report.msg.close_blocked", U("\u51e6\u7406\u4e2d\u306f\u9589\u3058\u3089\u308c\u307e\u305b\u3093\u3002")
    d.Add "report.msg.open_failed", U("\u958b\u3051\u307e\u305b\u3093: \u30d5\u30a1\u30a4\u30eb\u304c\u898b\u3064\u304b\u3089\u306a\u3044\u304b\u3001\u64cd\u4f5c\u304c\u30d6\u30ed\u30c3\u30af\u3055\u308c\u3066\u3044\u307e\u3059\u3002")
    d.Add "report.msg.select_row", U("\u8a73\u7d30\u3092\u8868\u793a\u3059\u308b\u306b\u306f\u884c\u3092\u9078\u629e \u00b7 PDF \u3092\u958b\u304f\u306b\u306f\u30c0\u30d6\u30eb\u30af\u30ea\u30c3\u30af")
    d.Add "report.msg.retry.none", U("\u518d\u8a66\u884c: \u30a8\u30e9\u30fc\u72b6\u614b\u306e\u30d5\u30a1\u30a4\u30eb\u306f\u3042\u308a\u307e\u305b\u3093\u3002")
    d.Add "report.msg.retry.error", U("\u518d\u8a66\u884c\u30a8\u30e9\u30fc: {0}")
    d.Add "report.msg.retry.partial", U("\u518d\u8a66\u884c: {0} \u6210\u529f\u3001{1} \u306f\u307e\u3060\u30a8\u30e9\u30fc\u3002")
    d.Add "report.msg.retry.all_fixed", U("\u518d\u8a66\u884c: {0} \u6210\u529f\u3001\u3059\u3079\u3066\u4fee\u6b63\u3055\u308c\u307e\u3057\u305f\u3002")
    d.Add "report.msg.copied", U("\u30b3\u30d4\u30fc\u3057\u307e\u3057\u305f!")
    d.Add "report.msg.exported", U("\u30ec\u30dd\u30fc\u30c8\u3092\u30e1\u30e2\u5e33\u3067\u958b\u304d\u307e\u3057\u305f\u3002")
    d.Add "report.msg.export_failed", U("\u30ec\u30dd\u30fc\u30c8\u3092\u4fdd\u5b58\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002")
    d.Add "reset.msg.all_defaults", U("\u3059\u3079\u3066\u306e\u8a2d\u5b9a\u3092\u65e2\u5b9a\u306b\u623b\u3057\u307e\u3057\u305f\u3002")
    d.Add "reset.lbl.registry", U("\u30ec\u30b8\u30b9\u30c8\u30ea")
    d.Add "reset.lbl.sandbox", U("\u30b5\u30f3\u30c9\u30dc\u30c3\u30af\u30b9")
    d.Add "reset.lbl.state", U("\u72b6\u614b")
    d.Add "reset.state.deleted", U("\u524a\u9664\u6e08\u307f")
    d.Add "reset.state.reinitialized", U("\u518d\u521d\u671f\u5316\u6e08\u307f")
    d.Add "reset.msg.failed", U("\u30ea\u30bb\u30c3\u30c8\u306b\u5931\u6557\u3057\u307e\u3057\u305f")
    d.Add "reset.msg.registry_complete", U("\u30ec\u30b8\u30b9\u30c8\u30ea\u3092\u30ea\u30bb\u30c3\u30c8\u3057\u307e\u3057\u305f\u3002")
    d.Add "reset.msg.sandbox_complete", U("\u30b5\u30f3\u30c9\u30dc\u30c3\u30af\u30b9\u3092\u30ea\u30bb\u30c3\u30c8\u3057\u307e\u3057\u305f")
    d.Add "scan.badge.tooltip", U("\u5408\u8a08: {0} | \u6700\u65b0: {1} | \u6b20\u843d: {2} | \u53e4\u3044: {3} | \u4e0d\u4e00\u81f4: {4} | \u30a8\u30e9\u30fc: {5}")
    d.Add "report.msg.opening", U("IDW \u30d5\u30a1\u30a4\u30eb\u3092\u958b\u3044\u3066\u3044\u307e\u3059...")
    d.Add "report.msg.updating", U("\u30d3\u30e5\u30fc\u3092\u66f4\u65b0\u4e2d...")
    d.Add "report.msg.exporting", U("PDF \u3092\u30a8\u30af\u30b9\u30dd\u30fc\u30c8\u4e2d...")
    d.Add "msg.no.write.access", U("\u30b5\u30f3\u30c9\u30dc\u30c3\u30af\u30b9\u30d5\u30a9\u30eb\u30c0\u30fc\u306b\u66f8\u304d\u8fbc\u3081\u307e\u305b\u3093\u3002\u6a29\u9650\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044:")
    d.Add "msg.pdf.temp_missing", U("\u4e00\u6642 PDF \u304c\u30a8\u30af\u30b9\u30dd\u30fc\u30c8\u4e2d\u306b\u6d88\u3048\u307e\u3057\u305f (\u30a6\u30a4\u30eb\u30b9\u5bfe\u7b56\u306e\u9694\u96e2\u307e\u305f\u306f\u30d5\u30a1\u30a4\u30eb \u30ed\u30c3\u30af\u306e\u53ef\u80fd\u6027): {0}")
    d.Add "msg.batch.abort.network", U("\u30d0\u30c3\u30c1\u4e2d\u6b62: {0} \u4ef6\u9023\u7d9a\u306e\u30cd\u30c3\u30c8\u30ef\u30fc\u30af\u30a8\u30e9\u30fc\u3002\u63a5\u7d9a\u3092\u78ba\u8a8d\u3057\u3066\u518d\u8a66\u884c\u3057\u3066\u304f\u3060\u3055\u3044\u3002")
    d.Add "msg.already.running", U("Batch PDF Publisher \u306f\u65e2\u306b\u5b9f\u884c\u4e2d\u3067\u3059\u30022 \u3064\u76ee\u306e\u30bb\u30c3\u30b7\u30e7\u30f3\u3092\u9589\u3058\u3066\u518d\u8a66\u884c\u3057\u3066\u304f\u3060\u3055\u3044\u3002")
    d.Add "report.search.placeholder", U("\u30d1\u30b9\u3092\u691c\u7d22...")
    d.Add "report.search.label", U("\u691c\u7d22:")
    d.Add "scan.collecting", U("\u30d5\u30a1\u30a4\u30eb\u3092\u53ce\u96c6\u4e2d...")
    d.Add "scan.analysing", U("\u5206\u6790: {0} / {1}...")
    d.Add "scan.auditing.start", U("\u30e2\u30c7\u30eb\u3092\u78ba\u8a8d\u4e2d...")
    d.Add "scan.auditing", U("\u76e3\u67fb: {0} / {1}...")
    d.Add "scan.cancelled", U("\u30ad\u30e3\u30f3\u30bb\u30eb\u6e08\u307f")
    d.Add "tip.search", U("\u30d1\u30b9\u3067\u30ec\u30dd\u30fc\u30c8\u884c\u3092\u30d5\u30a3\u30eb\u30bf\u30fc")
    d.Add "msg.confirm.reset", U("\u3059\u3079\u3066\u306e\u30aa\u30d7\u30b7\u30e7\u30f3\u3092\u65e2\u5b9a\u306b\u623b\u3057\u307e\u3059\u304b\uff1f")
    d.Add "scan.disk.insufficient", U("\u30c7\u30a3\u30b9\u30af\u5bb9\u91cf\u4e0d\u8db3\u3002\u7a7a\u304d: {0}")
    d.Add "scan.title.no_folder", U("\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u9078\u629e")
    d.Add "summary.hdr.models_full", U("{0} \u30e2\u30c7\u30eb ({1} \u30a2\u30bb\u30f3\u30d6\u30ea\u3001{2} \u30d1\u30fc\u30c4)")
    d.Add "summary.hdr.with_idw", U("{0} \u56f3\u9762\u3042\u308a")
    d.Add "summary.hdr.no_idw", U("{0} \u56f3\u9762\u306a\u3057")
    d.Add "errdlg.label.context", U("\u30b3\u30f3\u30c6\u30ad\u30b9\u30c8")
    d.Add "errdlg.label.time", U("\u6642\u9593")
    d.Add "errdlg.label.log", U("\u30ed\u30b0")
    d.Add "errdlg.label.stack", U("\u30b9\u30bf\u30c3\u30af")
    d.Add "pdf.err.prepare", U("\u6e96\u5099\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {0}")
    d.Add "pdf.err.header", U("\u30a8\u30af\u30b9\u30dd\u30fc\u30c8\u3055\u308c\u305f\u4e00\u6642 PDF \u306f\u30d8\u30c3\u30c0\u30fc\u691c\u8a3c\u306b\u5408\u683c\u3057\u307e\u305b\u3093\u3067\u3057\u305f")
    d.Add "pdf.err.unknown", U("\u4e0d\u660e\u306a\u30a8\u30e9\u30fc")
    d.Add "report.filter.all", U("\u3059\u3079\u3066")
    d.Add "report.filter.err", U("\u30a8\u30e9\u30fc")
    d.Add "report.filter.skip", U("\u30b9\u30ad\u30c3\u30d7")
    d.Add "report.filter.ok", U("OK")
    d.Add "report.filter.noidw", U(".idw \u306a\u3057")
    d.Add "scan.reset.sure", U("\u78ba\u5b9f\u3067\u3059\u304b\uff1f")
    d.Add "scan.reset.done", U("\u5b8c\u4e86")
    d.Add "report.form.title", U("\u51e6\u7406\u30ec\u30dd\u30fc\u30c8")
    d.Add "report.hdr.sub", U("\u30b5\u30d6\u30d5\u30a9\u30eb\u30c0\u30fc")
    d.Add "report.hdr.force_short", U("\u5f37\u5236")
    d.Add "report.hdr.audit", U("\u76e3\u67fb")
    d.Add "report.summary.full", U("{0} \u4f5c\u6210 \u00b7 {1} \u4e0a\u66f8\u304d \u00b7 {2} \u30b9\u30ad\u30c3\u30d7 \u00b7 {3} \u30a8\u30e9\u30fc \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} \u7a7a)")
    d.Add "scan.browse_title", U("\u30bd\u30fc\u30b9\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u9078\u629e")
End Sub

Private Sub L_FillKO(d As Object)
    d.Add "title", U("PDF \uc77c\uad04 \uac8c\uc2dc")
    d.Add "btn.run", U("\uc2e4\ud589")
    d.Add "btn.cancel", U("\ucde8\uc18c")
    d.Add "btn.browse", U("\ucc3e\uc544\ubcf4\uae30...")
    d.Add "btn.copy", U("\ubcf5\uc0ac")
    d.Add "btn.help", U("\ub3c4\uc6c0\ub9d0")
    d.Add "btn.settings", U("\uc124\uc815")
    d.Add "btn.back", U("\ub4a4\ub85c")
    d.Add "btn.export", U("\ub0b4\ubcf4\ub0b4\uae30")
    d.Add "btn.retry", U("\ub2e4\uc2dc \uc2dc\ub3c4")
    d.Add "btn.reset", U("\ucd08\uae30\ud654")
    d.Add "tab.scan", U("\uac80\uc0ac")
    d.Add "tab.settings", U("\uc124\uc815")
    d.Add "tab.log", U("\ub85c\uadf8")
    d.Add "settings.lang", U("\uc5b8\uc5b4:")
    d.Add "settings.lang.auto", U("\uc790\ub3d9 (\ud638\uc2a4\ud2b8)")
    d.Add "settings.loglevel", U("\ub85c\uadf8 \uc218\uc900:")
    d.Add "settings.loglevel.debug", U("\ub514\ubc84\uadf8")
    d.Add "settings.loglevel.info", U("\uc815\ubcf4")
    d.Add "settings.loglevel.warn", U("\uacbd\uace0")
    d.Add "settings.loglevel.error", U("\uc624\ub958\ub9cc")
    d.Add "lbl.folder", U("\uc6d0\ubcf8 \ud3f4\ub354:")
    d.Add "lbl.subfolders", U("\ud558\uc704 \ud3f4\ub354 \ud3ec\ud568")
    d.Add "chk.force", U("\uac15\uc81c \uc7ac\uc0dd\uc131")
    d.Add "chk.audit", U("\ub3c4\uba74 \uac80\uc0ac")
    d.Add "lbl.options", U("\ucd9c\ub825 \uc635\uc158:")
    d.Add "opt.single", U("\ub2e8\uc77c PDF (\ud1b5\ud569)")
    d.Add "opt.multi", U("\uac1c\ubcc4 \ud30c\uc77c")
    d.Add "chk.open", U("\uc644\ub8cc \ud6c4 \ud3f4\ub354 \uc5f4\uae30")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", U("\ubc1d\uac8c")
    d.Add "theme.dark", U("\uc5b4\ub461\uac8c")
    d.Add "theme.system", U("\uc2dc\uc2a4\ud15c")
    d.Add "msg.select.folder", U("\uc2a4\uce94\ud560 \ud3f4\ub354\ub97c \uc120\ud0dd\ud558\uc138\uc694.")
    d.Add "msg.no.idw", U("Inventor \ub3c4\uba74 (.idw) \uc744 \ucc3e\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4.")
    d.Add "msg.confirm.start", U("{0} \uac1c \ud30c\uc77c\uc744 \ucc98\ub9ac\ud558\uc2dc\uaca0\uc2b5\ub2c8\uae4c?")
    d.Add "msg.process.done", U("\ucc98\ub9ac\uac00 \uc131\uacf5\uc801\uc73c\ub85c \uc644\ub8cc\ub418\uc5c8\uc2b5\ub2c8\ub2e4.")
    d.Add "msg.error.title", U("\uc2dc\uc2a4\ud15c \uc624\ub958")
    d.Add "log.start", U("\uc138\uc158 \uc2dc\uc791: {0}")
    d.Add "log.end", U("\uc138\uc158 \uc885\ub8cc: {0}")
    d.Add "log.save", U("\ub85c\uadf8 \uc800\uc7a5: {0}")
    d.Add "scan.sum.drawings", U("{0} \ub3c4\uba74")
    d.Add "scan.sum.subfolders", U("\ud558\uc704 \ud3f4\ub354 \ud3ec\ud568")
    d.Add "scan.sum.ok", U("{0} \ucd5c\uc2e0")
    d.Add "scan.sum.all.current", U("\ubaa8\ub4e0 \ucd5c\uc2e0")
    d.Add "scan.sum.no.idw", U("\ub3c4\uba74\uc744 \ucc3e\uc744 \uc218 \uc5c6\uc74c")
    d.Add "scan.sum.missing", U("{0} \ub204\ub77d")
    d.Add "scan.sum.stale", U("{0} \uc624\ub798\ub428")
    d.Add "scan.sum.mismatch", U("{0} \ubd88\uc77c\uce58")
    d.Add "scan.sum.to_create", U("{0} \uc0dd\uc131")
    d.Add "scan.sum.to_recreate", U("{0} \uc7ac\uc0dd\uc131")
    d.Add "scan.sum.force_suffix", U("(\uac15\uc81c)")
    d.Add "scan.validate.error", U("\ud3f4\ub354 \uc720\ud6a8\uc131 \uac80\uc0ac \uc624\ub958.")
    d.Add "scan.sum.errors", U("{0} \uc624\ub958")
    d.Add "scan.no.folder", U("\ud3f4\ub354\uac00 \uc120\ud0dd\ub418\uc9c0 \uc54a\uc558\uc2b5\ub2c8\ub2e4.")
    d.Add "scan.no.folder.exists", U("\ud3f4\ub354\uac00 \uc874\uc7ac\ud558\uc9c0 \uc54a\uc2b5\ub2c8\ub2e4.")
    d.Add "scan.path.ok", U("\ud3f4\ub354\ub97c \ucc3e\uc558\uc2b5\ub2c8\ub2e4 \u2014 \uc2a4\uce94 \uc911...")
    d.Add "scan.path.too.long", U("\uacbd\ub85c\uac00 {0} \uc790 \uc81c\ud55c\uc744 \ucd08\uacfc\ud588\uc2b5\ub2c8\ub2e4.")
    d.Add "progress.processing", U("\ucc98\ub9ac \uc911...")
    d.Add "progress.init", U("\ucd08\uae30\ud654 \uc911...")
    d.Add "progress.cancel", U("\ucde8\uc18c")
    d.Add "progress.cancelling", U("\ucde8\uc18c \uc911...")
    d.Add "scan.models.detail", U("\ub3c4\uba74 \uc788\uc74c: {0} | \uc5c6\uc74c: {1}")
    d.Add "status.idle", U("\uc900\ube44\ub428")
    d.Add "status.scanning", U("\ud3f4\ub354 \uc2a4\uce94 \uc911...")
    d.Add "status.publishing", U("\uac8c\uc2dc \uc911: {0} / {1}...")
    d.Add "pdf.err.open", U("IDW \ud30c\uc77c\uc744 \uc5f4 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4.")
    d.Add "pdf.err.pagematch", U("PDF\ub294 {0} \ud398\uc774\uc9c0, {1} \ud398\uc774\uc9c0\uac00 \ud544\uc694\ud569\ub2c8\ub2e4.")
    d.Add "pdf.err.noext", U("\uc798\ubabb\ub41c \uacbd\ub85c (\ud655\uc7a5\uc790 \uc5c6\uc74c): {0}")
    d.Add "pdf.skipped.exists", U("\uac74\ub108\ub700 (\uc788\uc74c)")
    d.Add "pdf.skipped.exists.pages", U("\uac74\ub108\ub700 (\uc788\uc74c) \u00b7 {0}/{1} \ud398\uc774\uc9c0")
    d.Add "pdf.err.addin", U("PDF Translator \ucd94\uac00 \uae30\ub2a5\uc744 \ucc3e\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4.")
    d.Add "pdf.err.nofile", U("PDF \ub0b4\ubcf4\ub0b4\uae30\uc5d0\uc11c \ud30c\uc77c\uc774 \uc0dd\uc131\ub418\uc9c0 \uc54a\uc558\uc2b5\ub2c8\ub2e4.")
    d.Add "pdf.err.export", U("\ub0b4\ubcf4\ub0b4\uae30 \uc2e4\ud328: {0}")
    d.Add "pdf.err.bak.delete", U("\uc624\ub798\ub41c \ubc31\uc5c5\uc744 \uc0ad\uc81c\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4: {0}")
    d.Add "pdf.err.bak.move", U("\uae30\uc874 PDF\ub97c \ubc31\uc5c5\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4: {0}")
    d.Add "pdf.err.write", U("PDF\ub97c \uc4f8 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4: {0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 \ub0b4\ubcf4\ub0b4\uae30 \ubcf4\uace0\uc11c")
    d.Add "report.hdr.mode", U("\ubaa8\ub4dc: {0}")
    d.Add "report.hdr.folder", U("\ud3f4\ub354: {0}")
    d.Add "report.hdr.subfolders", U("\ud558\uc704 \ud3f4\ub354")
    d.Add "report.hdr.force", U("\uac15\uc81c \ub36e\uc5b4\uc4f0\uae30")
    d.Add "report.hdr.generated", U("\uc0dd\uc131: {0}")
    d.Add "report.state.on", U("\ucf1c\uc9d0")
    d.Add "report.state.off", U("\uaebc\uc9d0")
    d.Add "report.hdr.pdf.section", U("PDF \ub0b4\ubcf4\ub0b4\uae30")
    d.Add "report.hdr.pdf.detail", U("\uc138\ubd80 \uc815\ubcf4:")
    d.Add "report.created", U("\uc0dd\uc131\ub428")
    d.Add "report.overwritten", U("\ub36e\uc5b4\uc4f0\uae30\ub428")
    d.Add "report.skipped", U("\uac74\ub108\ub700")
    d.Add "report.errors", U("\uc624\ub958")
    d.Add "report.time", U("\uc2dc\uac04")
    d.Add "report.empty", U("\ube44\uc5b4 \uc788\uc74c")
    d.Add "report.pages", U("\ud398\uc774\uc9c0")
    d.Add "report.no.views", U("\ubdf0 \uc5c6\uc74c")
    d.Add "report.tag.created", U("[\uc0dd\uc131\ub428]")
    d.Add "report.tag.overwritten", U("[\ub36e\uc5b4\uc4f0\uae30\ub428]")
    d.Add "report.tag.skipped", U("[\uac74\ub108\ub700]")
    d.Add "report.tag.error", U("[\uc624\ub958]")
    d.Add "report.tag.hasdrawing", U("[\ub3c4\uba74 \uc788\uc74c]")
    d.Add "report.tag.nodrawing", U("[\ub3c4\uba74 \uc5c6\uc74c]")
    d.Add "report.hdr.audit.section", U("\ub3c4\uba74 \uac10\uc0ac")
    d.Add "report.hdr.audit.detail", U("\uc138\ubd80 \uc815\ubcf4:")
    d.Add "report.audit.title", U("\ub3c4\uba74 \uac10\uc0ac")
    d.Add "errdlg.retry.prompt", U("\ub2e4\uc2dc \uc2dc\ub3c4 = \ub2e4\uc2dc \uc2dc\ub3c4 | \ubb34\uc2dc = \uac74\ub108\ub6f0\uae30 | \uc911\ub2e8 = \uc911\uc9c0")
    d.Add "tip.run", U("\uc77c\uad04 PDF \ub0b4\ubcf4\ub0b4\uae30 \uc2dc\uc791")
    d.Add "tip.theme", U("\ud14c\ub9c8 \uc21c\ud658 (Inventor / \uc2dc\uc2a4\ud15c / \ub77c\uc774\ud2b8 / \ub2e4\ud06c)")
    d.Add "tip.reset", U("\ubaa8\ub4e0 \uc635\uc158\uc744 \uae30\ubcf8\uac12\uc73c\ub85c \uc7ac\uc124\uc815")
    d.Add "tip.cancel", U("\ud604\uc7ac \uc791\uc5c5 \ucde8\uc18c (ESC)")
    d.Add "tip.filter.nodrawing", U("\ub3c4\uba74 \uc5c6\ub294 \ubaa8\ub378\ub9cc \ud45c\uc2dc")
    d.Add "tip.list", U("\ub450 \ubc88 \ud074\ub9ad\ud558\uc5ec PDF \uc5f4\uae30; \ud074\ub9ad\ud558\uc5ec \uc138\ubd80 \uc815\ubcf4 \ubcf4\uae30")
    d.Add "tip.subfolders", U("\ud558\uc704 \ud3f4\ub354\ub97c \uc7ac\uadc0\uc801\uc73c\ub85c \uc2a4\uce94 (OldVersions \uc81c\uc678)")
    d.Add "tip.force", U("IDW\ubcf4\ub2e4 \ucd5c\uc2e0\uc774\uc5b4\ub3c4 PDF \uc7ac\uc0dd\uc131")
    d.Add "tip.audit", U("\ubaa8\ub378 (.ipt/.iam) \uc758 \ub204\ub77d\ub41c \ub3c4\uba74 \ud655\uc778")
    d.Add "tip.retry", U("\uc624\ub958\uac00 \uc788\ub294 \ubaa8\ub4e0 \ud30c\uc77c \uc7ac\uc2dc\ub3c4")
    d.Add "tip.export", U("\ubcf4\uace0\uc11c\ub97c \ud30c\uc77c\ub85c \uc800\uc7a5\ud558\uace0 \uba54\ubaa8\uc7a5\uc5d0\uc11c \uc5f4\uae30")
    d.Add "report.msg.close_blocked", U("\ucc98\ub9ac \uc911\uc5d0\ub294 \ub2eb\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4.")
    d.Add "report.msg.open_failed", U("\uc5f4 \uc218 \uc5c6\uc74c: \ud30c\uc77c\uc744 \ucc3e\uc744 \uc218 \uc5c6\uac70\ub098 \uc791\uc5c5\uc774 \ucc28\ub2e8\ub418\uc5c8\uc2b5\ub2c8\ub2e4.")
    d.Add "report.msg.select_row", U("\uc138\ubd80 \uc815\ubcf4\ub97c \ubcf4\ub824\uba74 \ud589 \uc120\ud0dd \u00b7 PDF\ub97c \uc5f4\ub824\uba74 \ub450 \ubc88 \ud074\ub9ad")
    d.Add "report.msg.retry.none", U("\ub2e4\uc2dc \uc2dc\ub3c4: \uc624\ub958 \uc0c1\ud0dc\uc758 \ud30c\uc77c\uc774 \uc5c6\uc2b5\ub2c8\ub2e4.")
    d.Add "report.msg.retry.error", U("\ub2e4\uc2dc \uc2dc\ub3c4 \uc624\ub958: {0}")
    d.Add "report.msg.retry.partial", U("\ub2e4\uc2dc \uc2dc\ub3c4: {0} \uc131\uacf5, {1} \uc5ec\uc804\ud788 \uc624\ub958.")
    d.Add "report.msg.retry.all_fixed", U("\ub2e4\uc2dc \uc2dc\ub3c4: {0} \uc131\uacf5, \ubaa8\ub450 \uc218\uc815\ub418\uc5c8\uc2b5\ub2c8\ub2e4.")
    d.Add "report.msg.copied", U("\ubcf5\uc0ac\ub428!")
    d.Add "report.msg.exported", U("\ubcf4\uace0\uc11c\ub97c \uba54\ubaa8\uc7a5\uc5d0\uc11c \uc5f4\uc5c8\uc2b5\ub2c8\ub2e4.")
    d.Add "report.msg.export_failed", U("\ubcf4\uace0\uc11c\ub97c \uc800\uc7a5\ud558\uc9c0 \ubabb\ud588\uc2b5\ub2c8\ub2e4.")
    d.Add "reset.msg.all_defaults", U("\ubaa8\ub4e0 \uc124\uc815\uc774 \uae30\ubcf8\uac12\uc73c\ub85c \ubcf5\uc6d0\ub418\uc5c8\uc2b5\ub2c8\ub2e4.")
    d.Add "reset.lbl.registry", U("\ub808\uc9c0\uc2a4\ud2b8\ub9ac")
    d.Add "reset.lbl.sandbox", U("\uc0cc\ub4dc\ubc15\uc2a4")
    d.Add "reset.lbl.state", U("\uc0c1\ud0dc")
    d.Add "reset.state.deleted", U("\uc0ad\uc81c\ub428")
    d.Add "reset.state.reinitialized", U("\uc7ac\ucd08\uae30\ud654\ub428")
    d.Add "reset.msg.failed", U("\uc7ac\uc124\uc815 \uc2e4\ud328")
    d.Add "reset.msg.registry_complete", U("\ub808\uc9c0\uc2a4\ud2b8\ub9ac\uac00 \uc7ac\uc124\uc815\ub418\uc5c8\uc2b5\ub2c8\ub2e4.")
    d.Add "reset.msg.sandbox_complete", U("\uc0cc\ub4dc\ubc15\uc2a4\uac00 \uc7ac\uc124\uc815\ub418\uc5c8\uc2b5\ub2c8\ub2e4.")
    d.Add "scan.badge.tooltip", U("\ud569\uacc4: {0} | \ucd5c\uc2e0: {1} | \ub204\ub77d: {2} | \uc624\ub798\ub428: {3} | \ubd88\uc77c\uce58: {4} | \uc624\ub958: {5}")
    d.Add "report.msg.opening", U("IDW \ud30c\uc77c \uc5ec\ub294 \uc911...")
    d.Add "report.msg.updating", U("\ubdf0 \uc5c5\ub370\uc774\ud2b8 \uc911...")
    d.Add "report.msg.exporting", U("PDF \ub0b4\ubcf4\ub0b4\ub294 \uc911...")
    d.Add "msg.no.write.access", U("\uc0cc\ub4dc\ubc15\uc2a4 \ud3f4\ub354\uc5d0 \uc4f8 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4. \uad8c\ud55c\uc744 \ud655\uc778\ud558\uc138\uc694:")
    d.Add "msg.pdf.temp_missing", U("\ub0b4\ubcf4\ub0b4\uae30 \uc911 \uc784\uc2dc PDF\uac00 \uc0ac\ub77c\uc84c\uc2b5\ub2c8\ub2e4 (\ubc14\uc774\ub7ec\uc2a4 \ubc31\uc2e0 \uacbd\ub9ac \ub610\ub294 \ud30c\uc77c \uc7a0\uae08 \uac00\ub2a5\uc131): {0}")
    d.Add "msg.batch.abort.network", U("\uc77c\uad04 \ucc98\ub9ac \uc911\ub2e8: {0}\ud68c \uc5f0\uc18d \ub124\ud2b8\uc6cc\ud06c \uc624\ub958. \uc5f0\uacb0\uc744 \ud655\uc778\ud558\uace0 \ub2e4\uc2dc \uc2dc\ub3c4\ud558\uc138\uc694.")
    d.Add "msg.already.running", U("Batch PDF Publisher\uac00 \uc774\ubbf8 \uc2e4\ud589 \uc911\uc785\ub2c8\ub2e4. \ub450 \ubc88\uc9f8 \uc138\uc158\uc744 \ub2eb\uace0 \ub2e4\uc2dc \uc2dc\ub3c4\ud558\uc138\uc694.")
    d.Add "report.search.placeholder", U("\uacbd\ub85c \uac80\uc0c9...")
    d.Add "report.search.label", U("\uac80\uc0c9:")
    d.Add "scan.collecting", U("\ud30c\uc77c \uc218\uc9d1 \uc911...")
    d.Add "scan.analysing", U("\ubd84\uc11d: {0} / {1}...")
    d.Add "scan.auditing.start", U("\ubaa8\ub378 \uac80\uc0ac \uc911...")
    d.Add "scan.auditing", U("\uac10\uc0ac: {0} / {1}...")
    d.Add "scan.cancelled", U("\ucde8\uc18c\ub428")
    d.Add "tip.search", U("\uacbd\ub85c\ub85c \ubcf4\uace0\uc11c \ud589 \ud544\ud130\ub9c1")
    d.Add "msg.confirm.reset", U("\ubaa8\ub4e0 \uc635\uc158\uc744 \uae30\ubcf8\uac12\uc73c\ub85c \uc7ac\uc124\uc815\ud558\uc2dc\uaca0\uc2b5\ub2c8\uae4c?")
    d.Add "scan.disk.insufficient", U("\ub514\uc2a4\ud06c \uacf5\uac04 \ubd80\uc871. \uc5ec\uc720: {0}")
    d.Add "scan.title.no_folder", U("\ud3f4\ub354 \uc120\ud0dd")
    d.Add "summary.hdr.models_full", U("{0}\uac1c \ubaa8\ub378 ({1}\uac1c \uc5b4\uc148\ube14\ub9ac, {2}\uac1c \ubd80\ud488)")
    d.Add "summary.hdr.with_idw", U("{0}\uac1c \ub3c4\uba74 \uc788\uc74c")
    d.Add "summary.hdr.no_idw", U("{0}\uac1c \ub3c4\uba74 \uc5c6\uc74c")
    d.Add "errdlg.label.context", U("\ucee8\ud14d\uc2a4\ud2b8")
    d.Add "errdlg.label.time", U("\uc2dc\uac04")
    d.Add "errdlg.label.log", U("\ub85c\uadf8")
    d.Add "errdlg.label.stack", U("\uc2a4\ud0dd")
    d.Add "pdf.err.prepare", U("\uc900\ube44 \uc2e4\ud328: {0}")
    d.Add "pdf.err.header", U("\ub0b4\ubcf4\ub0b8 \uc784\uc2dc PDF\uac00 \ud5e4\ub354 \uc720\ud6a8\uc131 \uac80\uc0ac\ub97c \ud1b5\uacfc\ud558\uc9c0 \ubabb\ud588\uc2b5\ub2c8\ub2e4")
    d.Add "pdf.err.unknown", U("\uc54c \uc218 \uc5c6\ub294 \uc624\ub958")
    d.Add "report.filter.all", U("\ubaa8\ub450")
    d.Add "report.filter.err", U("\uc624\ub958")
    d.Add "report.filter.skip", U("\uac74\ub108\ub700")
    d.Add "report.filter.ok", U("OK")
    d.Add "report.filter.noidw", U(".idw \uc5c6\uc74c")
    d.Add "scan.reset.sure", U("\ud655\uc2e4\ud569\ub2c8\uae4c?")
    d.Add "scan.reset.done", U("\uc644\ub8cc")
    d.Add "report.form.title", U("\ucc98\ub9ac \ubcf4\uace0\uc11c")
    d.Add "report.hdr.sub", U("\ud558\uc704 \ud3f4\ub354")
    d.Add "report.hdr.force_short", U("\uac15\uc81c")
    d.Add "report.hdr.audit", U("\uac10\uc0ac")
    d.Add "report.summary.full", U("{0} \uc0dd\uc131 \u00b7 {1} \ub36e\uc5b4\uc4f0\uae30 \u00b7 {2} \uac74\ub108\ub700 \u00b7 {3} \uc624\ub958 \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} \ube44\uc5b4 \uc788\uc74c)")
    d.Add "scan.browse_title", U("\uc6d0\ubcf8 \ud3f4\ub354 \uc120\ud0dd")
End Sub

Private Sub L_FillZH(d As Object)
    d.Add "title", U("PDF \u6279\u91cf\u53d1\u5e03")
    d.Add "btn.run", U("\u8fd0\u884c")
    d.Add "btn.cancel", U("\u53d6\u6d88")
    d.Add "btn.browse", U("\u6d4f\u89c8...")
    d.Add "btn.copy", U("\u590d\u5236")
    d.Add "btn.help", U("\u5e2e\u52a9")
    d.Add "btn.settings", U("\u8bbe\u7f6e")
    d.Add "btn.back", U("\u8fd4\u56de")
    d.Add "btn.export", U("\u5bfc\u51fa")
    d.Add "btn.retry", U("\u91cd\u8bd5")
    d.Add "btn.reset", U("\u91cd\u7f6e")
    d.Add "tab.scan", U("\u626b\u63cf")
    d.Add "tab.settings", U("\u8bbe\u7f6e")
    d.Add "tab.log", U("\u65e5\u5fd7")
    d.Add "settings.lang", U("\u8bed\u8a00\uff1a")
    d.Add "settings.lang.auto", U("\u81ea\u52a8\uff08\u4e3b\u673a\uff09")
    d.Add "settings.loglevel", U("\u65e5\u5fd7\u7ea7\u522b\uff1a")
    d.Add "settings.loglevel.debug", U("\u8c03\u8bd5")
    d.Add "settings.loglevel.info", U("\u4fe1\u606f")
    d.Add "settings.loglevel.warn", U("\u8b66\u544a")
    d.Add "settings.loglevel.error", U("\u4ec5\u9519\u8bef")
    d.Add "lbl.folder", U("\u6e90\u6587\u4ef6\u5939\uff1a")
    d.Add "lbl.subfolders", U("\u5305\u62ec\u5b50\u6587\u4ef6\u5939")
    d.Add "chk.force", U("\u5f3a\u5236\u91cd\u65b0\u751f\u6210")
    d.Add "chk.audit", U("\u56fe\u7eb8\u5ba1\u8ba1")
    d.Add "lbl.options", U("\u8f93\u51fa\u9009\u9879\uff1a")
    d.Add "opt.single", U("\u5355\u4e2a PDF\uff08\u5408\u5e76\uff09")
    d.Add "opt.multi", U("\u72ec\u7acb\u6587\u4ef6")
    d.Add "chk.open", U("\u5b8c\u6210\u540e\u6253\u5f00\u6587\u4ef6\u5939")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", U("\u6d45\u8272")
    d.Add "theme.dark", U("\u6df1\u8272")
    d.Add "theme.system", U("\u7cfb\u7edf")
    d.Add "msg.select.folder", U("\u8bf7\u9009\u62e9\u8981\u626b\u63cf\u7684\u6587\u4ef6\u5939\u3002")
    d.Add "msg.no.idw", U("\u672a\u627e\u5230 Inventor \u56fe\u7eb8 (.idw)\u3002")
    d.Add "msg.confirm.start", U("\u662f\u5426\u5f00\u59cb\u5904\u7406 {0} \u4e2a\u6587\u4ef6\uff1f")
    d.Add "msg.process.done", U("\u5904\u7406\u5df2\u6210\u529f\u5b8c\u6210\u3002")
    d.Add "msg.error.title", U("\u7cfb\u7edf\u9519\u8bef")
    d.Add "log.start", U("\u4f1a\u8bdd\u5f00\u59cb\uff1a{0}")
    d.Add "log.end", U("\u4f1a\u8bdd\u7ed3\u675f\uff1a{0}")
    d.Add "log.save", U("\u65e5\u5fd7\u5df2\u4fdd\u5b58\u5230\uff1a{0}")
    d.Add "scan.sum.drawings", U("{0} \u4e2a\u56fe\u7eb8")
    d.Add "scan.sum.subfolders", U("\u5305\u62ec\u5b50\u6587\u4ef6\u5939")
    d.Add "scan.sum.ok", U("{0} \u4e2a\u6700\u65b0")
    d.Add "scan.sum.all.current", U("\u5168\u90e8\u6700\u65b0")
    d.Add "scan.sum.no.idw", U("\u672a\u627e\u5230\u56fe\u7eb8")
    d.Add "scan.sum.missing", U("{0} \u4e2a\u7f3a\u5931")
    d.Add "scan.sum.stale", U("{0} \u4e2a\u8fc7\u65f6")
    d.Add "scan.sum.mismatch", U("{0} \u4e2a\u4e0d\u5339\u914d")
    d.Add "scan.sum.to_create", U("{0} \u521b\u5efa")
    d.Add "scan.sum.to_recreate", U("{0} \u91cd\u5efa")
    d.Add "scan.sum.force_suffix", U("(\u5f3a\u5236)")
    d.Add "scan.validate.error", U("\u6587\u4ef6\u5939\u9a8c\u8bc1\u9519\u8bef\u3002")
    d.Add "scan.sum.errors", U("{0} \u4e2a\u9519\u8bef")
    d.Add "scan.no.folder", U("\u672a\u9009\u62e9\u6587\u4ef6\u5939\u3002")
    d.Add "scan.no.folder.exists", U("\u6587\u4ef6\u5939\u4e0d\u5b58\u5728\u3002")
    d.Add "scan.path.ok", U("\u5df2\u627e\u5230\u6587\u4ef6\u5939 \u2014 \u6b63\u5728\u626b\u63cf...")
    d.Add "scan.path.too.long", U("\u8def\u5f84\u8d85\u8fc7 {0} \u4e2a\u5b57\u7b26\u7684\u9650\u5236\u3002")
    d.Add "progress.processing", U("\u6b63\u5728\u5904\u7406...")
    d.Add "progress.init", U("\u6b63\u5728\u521d\u59cb\u5316...")
    d.Add "progress.cancel", U("\u53d6\u6d88")
    d.Add "progress.cancelling", U("\u6b63\u5728\u53d6\u6d88...")
    d.Add "scan.models.detail", U("\u6709\u56fe\u7eb8\uff1a{0} | \u65e0\uff1a{1}")
    d.Add "status.idle", U("\u5c31\u7eea")
    d.Add "status.scanning", U("\u6b63\u5728\u626b\u63cf\u6587\u4ef6\u5939...")
    d.Add "status.publishing", U("\u6b63\u5728\u53d1\u5e03\uff1a{0} / {1}...")
    d.Add "pdf.err.open", U("\u65e0\u6cd5\u6253\u5f00 IDW \u6587\u4ef6\u3002")
    d.Add "pdf.err.pagematch", U("PDF \u6709 {0} \u9875\uff0c\u9884\u671f {1} \u9875\u3002")
    d.Add "pdf.err.noext", U("\u8def\u5f84\u65e0\u6548\uff08\u65e0\u6269\u5c55\u540d\uff09\uff1a{0}")
    d.Add "pdf.skipped.exists", U("\u5df2\u8df3\u8fc7\uff08\u5b58\u5728\uff09")
    d.Add "pdf.skipped.exists.pages", U("\u5df2\u8df3\u8fc7\uff08\u5b58\u5728\uff09\u00b7 {0}/{1} \u9875")
    d.Add "pdf.err.addin", U("\u672a\u627e\u5230 PDF Translator \u52a0\u8f7d\u9879\u3002")
    d.Add "pdf.err.nofile", U("PDF \u5bfc\u51fa\u672a\u521b\u5efa\u6587\u4ef6\u3002")
    d.Add "pdf.err.export", U("\u5bfc\u51fa\u5931\u8d25\uff1a{0}")
    d.Add "pdf.err.bak.delete", U("\u65e0\u6cd5\u5220\u9664\u65e7\u5907\u4efd\uff1a{0}")
    d.Add "pdf.err.bak.move", U("\u65e0\u6cd5\u5907\u4efd\u73b0\u6709 PDF\uff1a{0}")
    d.Add "pdf.err.write", U("\u65e0\u6cd5\u5199\u5165 PDF\uff1a{0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 \u5bfc\u51fa\u62a5\u544a")
    d.Add "report.hdr.mode", U("\u6a21\u5f0f\uff1a{0}")
    d.Add "report.hdr.folder", U("\u6587\u4ef6\u5939\uff1a{0}")
    d.Add "report.hdr.subfolders", U("\u5b50\u6587\u4ef6\u5939")
    d.Add "report.hdr.force", U("\u5f3a\u5236\u8986\u76d6")
    d.Add "report.hdr.generated", U("\u751f\u6210\u65f6\u95f4\uff1a{0}")
    d.Add "report.state.on", U("\u5f00")
    d.Add "report.state.off", U("\u5173")
    d.Add "report.hdr.pdf.section", U("PDF \u5bfc\u51fa")
    d.Add "report.hdr.pdf.detail", U("\u8be6\u7ec6\u4fe1\u606f\uff1a")
    d.Add "report.created", U("\u5df2\u521b\u5efa")
    d.Add "report.overwritten", U("\u5df2\u8986\u76d6")
    d.Add "report.skipped", U("\u5df2\u8df3\u8fc7")
    d.Add "report.errors", U("\u9519\u8bef")
    d.Add "report.time", U("\u65f6\u95f4")
    d.Add "report.empty", U("\u7a7a")
    d.Add "report.pages", U("\u9875")
    d.Add "report.no.views", U("\u65e0\u89c6\u56fe")
    d.Add "report.tag.created", U("[\u5df2\u521b\u5efa]")
    d.Add "report.tag.overwritten", U("[\u5df2\u8986\u76d6]")
    d.Add "report.tag.skipped", U("[\u5df2\u8df3\u8fc7]")
    d.Add "report.tag.error", U("[\u9519\u8bef]")
    d.Add "report.tag.hasdrawing", U("[\u6709\u56fe\u7eb8]")
    d.Add "report.tag.nodrawing", U("[\u65e0\u56fe\u7eb8]")
    d.Add "report.hdr.audit.section", U("\u56fe\u7eb8\u5ba1\u8ba1")
    d.Add "report.hdr.audit.detail", U("\u8be6\u7ec6\u4fe1\u606f\uff1a")
    d.Add "report.audit.title", U("\u56fe\u7eb8\u5ba1\u8ba1")
    d.Add "errdlg.retry.prompt", U("\u91cd\u8bd5 = \u518d\u8bd5 | \u5ffd\u7565 = \u8df3\u8fc7 | \u4e2d\u6b62 = \u505c\u6b62")
    d.Add "tip.run", U("\u5f00\u59cb\u6279\u91cf PDF \u5bfc\u51fa")
    d.Add "tip.theme", U("\u5207\u6362\u4e3b\u9898\uff08Inventor / \u7cfb\u7edf / \u6d45\u8272 / \u6df1\u8272\uff09")
    d.Add "tip.reset", U("\u5c06\u6240\u6709\u9009\u9879\u6062\u590d\u4e3a\u9ed8\u8ba4\u503c")
    d.Add "tip.cancel", U("\u53d6\u6d88\u5f53\u524d\u64cd\u4f5c\uff08ESC\uff09")
    d.Add "tip.filter.nodrawing", U("\u4ec5\u663e\u793a\u65e0\u56fe\u7eb8\u7684\u6a21\u578b")
    d.Add "tip.list", U("\u53cc\u51fb\u6253\u5f00 PDF\uff1b\u5355\u51fb\u67e5\u770b\u8be6\u7ec6\u4fe1\u606f")
    d.Add "tip.subfolders", U("\u9012\u5f52\u626b\u63cf\u5b50\u6587\u4ef6\u5939\uff08\u8df3\u8fc7 OldVersions\uff09")
    d.Add "tip.force", U("\u5373\u4f7f PDF \u6bd4 IDW \u65b0\u4e5f\u91cd\u65b0\u521b\u5efa")
    d.Add "tip.audit", U("\u68c0\u67e5\u6a21\u578b (.ipt/.iam) \u662f\u5426\u7f3a\u5c11\u56fe\u7eb8")
    d.Add "tip.retry", U("\u91cd\u8bd5\u6240\u6709\u51fa\u9519\u7684\u6587\u4ef6")
    d.Add "tip.export", U("\u5c06\u62a5\u544a\u4fdd\u5b58\u5230\u6587\u4ef6\u5e76\u5728\u8bb0\u4e8b\u672c\u4e2d\u6253\u5f00")
    d.Add "report.msg.close_blocked", U("\u5904\u7406\u671f\u95f4\u65e0\u6cd5\u5173\u95ed\u3002")
    d.Add "report.msg.open_failed", U("\u65e0\u6cd5\u6253\u5f00\uff1a\u6587\u4ef6\u672a\u627e\u5230\u6216\u64cd\u4f5c\u88ab\u963b\u6b62\u3002")
    d.Add "report.msg.select_row", U("\u9009\u62e9\u884c\u4ee5\u67e5\u770b\u8be6\u7ec6\u4fe1\u606f \u00b7 \u53cc\u51fb\u6253\u5f00 PDF")
    d.Add "report.msg.retry.none", U("\u91cd\u8bd5\uff1a\u6ca1\u6709\u5904\u4e8e\u9519\u8bef\u72b6\u6001\u7684\u6587\u4ef6\u3002")
    d.Add "report.msg.retry.error", U("\u91cd\u8bd5\u9519\u8bef\uff1a{0}")
    d.Add "report.msg.retry.partial", U("\u91cd\u8bd5\uff1a{0} \u4e2a\u6210\u529f\uff0c{1} \u4e2a\u4ecd\u51fa\u9519\u3002")
    d.Add "report.msg.retry.all_fixed", U("\u91cd\u8bd5\uff1a{0} \u4e2a\u6210\u529f\uff0c\u5168\u90e8\u5df2\u4fee\u590d\u3002")
    d.Add "report.msg.copied", U("\u5df2\u590d\u5236\uff01")
    d.Add "report.msg.exported", U("\u62a5\u544a\u5df2\u5728\u8bb0\u4e8b\u672c\u4e2d\u6253\u5f00\u3002")
    d.Add "report.msg.export_failed", U("\u65e0\u6cd5\u4fdd\u5b58\u62a5\u544a\u3002")
    d.Add "reset.msg.all_defaults", U("\u6240\u6709\u8bbe\u7f6e\u5df2\u6062\u590d\u9ed8\u8ba4\u3002")
    d.Add "reset.lbl.registry", U("\u6ce8\u518c\u8868")
    d.Add "reset.lbl.sandbox", U("\u6c99\u76d2")
    d.Add "reset.lbl.state", U("\u72b6\u6001")
    d.Add "reset.state.deleted", U("\u5df2\u5220\u9664")
    d.Add "reset.state.reinitialized", U("\u5df2\u91cd\u65b0\u521d\u59cb\u5316")
    d.Add "reset.msg.failed", U("\u91cd\u7f6e\u5931\u8d25")
    d.Add "reset.msg.registry_complete", U("\u6ce8\u518c\u8868\u5df2\u91cd\u7f6e\u3002")
    d.Add "reset.msg.sandbox_complete", U("\u6c99\u76d2\u5df2\u91cd\u7f6e")
    d.Add "scan.badge.tooltip", U("\u603b\u8ba1\uff1a{0} | \u6700\u65b0\uff1a{1} | \u7f3a\u5931\uff1a{2} | \u8fc7\u65f6\uff1a{3} | \u4e0d\u5339\u914d\uff1a{4} | \u9519\u8bef\uff1a{5}")
    d.Add "report.msg.opening", U("\u6b63\u5728\u6253\u5f00 IDW \u6587\u4ef6...")
    d.Add "report.msg.updating", U("\u6b63\u5728\u66f4\u65b0\u89c6\u56fe...")
    d.Add "report.msg.exporting", U("\u6b63\u5728\u5bfc\u51fa PDF...")
    d.Add "msg.no.write.access", U("\u65e0\u6cd5\u5199\u5165\u6c99\u76d2\u6587\u4ef6\u5939\u3002\u8bf7\u68c0\u67e5\u6743\u9650\uff1a")
    d.Add "msg.pdf.temp_missing", U("\u5bfc\u51fa\u671f\u95f4\u4e34\u65f6 PDF \u6d88\u5931\uff08\u53ef\u80fd\u662f\u6740\u6bd2\u8f6f\u4ef6\u9694\u79bb\u6216\u6587\u4ef6\u9501\u5b9a\uff09\uff1a{0}")
    d.Add "msg.batch.abort.network", U("\u6279\u5904\u7406\u4e2d\u6b62\uff1a\u8fde\u7eed {0} \u4e2a\u7f51\u7edc\u9519\u8bef\u3002\u8bf7\u68c0\u67e5\u8fde\u63a5\u5e76\u91cd\u8bd5\u3002")
    d.Add "msg.already.running", U("Batch PDF Publisher \u5df2\u5728\u8fd0\u884c\u3002\u8bf7\u5173\u95ed\u7b2c\u4e8c\u4e2a\u4f1a\u8bdd\u5e76\u91cd\u8bd5\u3002")
    d.Add "report.search.placeholder", U("\u641c\u7d22\u8def\u5f84...")
    d.Add "report.search.label", U("\u641c\u7d22\uff1a")
    d.Add "scan.collecting", U("\u6b63\u5728\u6536\u96c6\u6587\u4ef6...")
    d.Add "scan.analysing", U("\u6b63\u5728\u5206\u6790\uff1a{0} / {1}...")
    d.Add "scan.auditing.start", U("\u6b63\u5728\u68c0\u67e5\u6a21\u578b...")
    d.Add "scan.auditing", U("\u5ba1\u8ba1\uff1a{0} / {1}...")
    d.Add "scan.cancelled", U("\u5df2\u53d6\u6d88")
    d.Add "tip.search", U("\u6309\u8def\u5f84\u7b5b\u9009\u62a5\u544a\u884c")
    d.Add "msg.confirm.reset", U("\u5c06\u6240\u6709\u9009\u9879\u6062\u590d\u4e3a\u9ed8\u8ba4\u503c\uff1f")
    d.Add "scan.disk.insufficient", U("\u78c1\u76d8\u7a7a\u95f4\u4e0d\u8db3\u3002\u53ef\u7528\uff1a{0}")
    d.Add "scan.title.no_folder", U("\u9009\u62e9\u6587\u4ef6\u5939")
    d.Add "summary.hdr.models_full", U("{0} \u4e2a\u6a21\u578b\uff08{1} \u4e2a\u88c5\u914d\uff0c{2} \u4e2a\u96f6\u4ef6\uff09")
    d.Add "summary.hdr.with_idw", U("{0} \u4e2a\u6709\u56fe\u7eb8")
    d.Add "summary.hdr.no_idw", U("{0} \u4e2a\u65e0\u56fe\u7eb8")
    d.Add "errdlg.label.context", U("\u4e0a\u4e0b\u6587")
    d.Add "errdlg.label.time", U("\u65f6\u95f4")
    d.Add "errdlg.label.log", U("\u65e5\u5fd7")
    d.Add "errdlg.label.stack", U("\u5806\u6808")
    d.Add "pdf.err.prepare", U("\u51c6\u5907\u5931\u8d25\uff1a{0}")
    d.Add "pdf.err.header", U("\u5bfc\u51fa\u7684\u4e34\u65f6 PDF \u672a\u901a\u8fc7\u6807\u5934\u9a8c\u8bc1")
    d.Add "pdf.err.unknown", U("\u672a\u77e5\u9519\u8bef")
    d.Add "report.filter.all", U("\u5168\u90e8")
    d.Add "report.filter.err", U("\u9519\u8bef")
    d.Add "report.filter.skip", U("\u5df2\u8df3\u8fc7")
    d.Add "report.filter.ok", U("OK")
    d.Add "report.filter.noidw", U("\u65e0 .idw")
    d.Add "scan.reset.sure", U("\u786e\u5b9a\uff1f")
    d.Add "scan.reset.done", U("\u5b8c\u6210")
    d.Add "report.form.title", U("\u5904\u7406\u62a5\u544a")
    d.Add "report.hdr.sub", U("\u5b50\u6587\u4ef6\u5939")
    d.Add "report.hdr.force_short", U("\u5f3a\u5236")
    d.Add "report.hdr.audit", U("\u5ba1\u8ba1")
    d.Add "report.summary.full", U("{0} \u5df2\u521b\u5efa \u00b7 {1} \u5df2\u8986\u76d6 \u00b7 {2} \u5df2\u8df3\u8fc7 \u00b7 {3} \u4e2a\u9519\u8bef \u00b7 {4}")
    d.Add "report.skip.empty", U("{0}\uff08{1} \u4e2a\u7a7a\uff09")
    d.Add "scan.browse_title", U("\u9009\u62e9\u6e90\u6587\u4ef6\u5939")
End Sub

Private Sub L_FillZHT(d As Object)
    d.Add "title", U("PDF \u6279\u6b21\u767c\u4f48")
    d.Add "btn.run", U("\u57f7\u884c")
    d.Add "btn.cancel", U("\u53d6\u6d88")
    d.Add "btn.browse", U("\u700f\u89bd...")
    d.Add "btn.copy", U("\u8907\u88fd")
    d.Add "btn.help", U("\u8aac\u660e")
    d.Add "btn.settings", U("\u8a2d\u5b9a")
    d.Add "btn.back", U("\u8fd4\u56de")
    d.Add "btn.export", U("\u532f\u51fa")
    d.Add "btn.retry", U("\u91cd\u8a66")
    d.Add "btn.reset", U("\u91cd\u8a2d")
    d.Add "tab.scan", U("\u6383\u63cf")
    d.Add "tab.settings", U("\u8a2d\u5b9a")
    d.Add "tab.log", U("\u8a18\u9304")
    d.Add "settings.lang", U("\u8a9e\u8a00\uff1a")
    d.Add "settings.lang.auto", U("\u81ea\u52d5 (\u4e3b\u6a5f)")
    d.Add "settings.loglevel", U("\u8a18\u9304\u5c64\u7d1a\uff1a")
    d.Add "settings.loglevel.debug", U("\u5075\u932f")
    d.Add "settings.loglevel.info", U("\u8cc7\u8a0a")
    d.Add "settings.loglevel.warn", U("\u8b66\u544a")
    d.Add "settings.loglevel.error", U("\u50c5\u932f\u8aa4")
    d.Add "lbl.folder", U("\u4f86\u6e90\u8cc7\u6599\u593e\uff1a")
    d.Add "lbl.subfolders", U("\u5305\u542b\u5b50\u8cc7\u6599\u593e")
    d.Add "chk.force", U("\u5f37\u5236\u91cd\u65b0\u7522\u751f")
    d.Add "chk.audit", U("\u5716\u9762\u7a3d\u6838")
    d.Add "lbl.options", U("\u8f38\u51fa\u9078\u9805\uff1a")
    d.Add "opt.single", U("\u55ae\u4e00 PDF (\u5408\u4f75)")
    d.Add "opt.multi", U("\u500b\u5225\u6a94\u6848")
    d.Add "chk.open", U("\u5b8c\u6210\u5f8c\u958b\u555f\u8cc7\u6599\u593e")
    d.Add "theme.inventor", "Inventor"
    d.Add "theme.light", U("\u6dfa\u8272")
    d.Add "theme.dark", U("\u6df1\u8272")
    d.Add "theme.system", U("\u7cfb\u7d71")
    d.Add "msg.select.folder", U("\u8acb\u9078\u64c7\u8981\u6383\u63cf\u7684\u8cc7\u6599\u593e\u3002")
    d.Add "msg.no.idw", U("\u627e\u4e0d\u5230 Inventor \u5716\u9762 (.idw)\u3002")
    d.Add "msg.confirm.start", U("\u662f\u5426\u958b\u59cb\u8655\u7406 {0} \u500b\u6a94\u6848\uff1f")
    d.Add "msg.process.done", U("\u8655\u7406\u5df2\u6210\u529f\u5b8c\u6210\u3002")
    d.Add "msg.error.title", U("\u7cfb\u7d71\u932f\u8aa4")
    d.Add "log.start", U("\u5de5\u4f5c\u968e\u6bb5\u958b\u59cb\uff1a{0}")
    d.Add "log.end", U("\u5de5\u4f5c\u968e\u6bb5\u7d50\u675f\uff1a{0}")
    d.Add "log.save", U("\u8a18\u9304\u5df2\u5132\u5b58\u81f3\uff1a{0}")
    d.Add "scan.sum.drawings", U("{0} \u500b\u5716\u9762")
    d.Add "scan.sum.subfolders", U("\u5305\u542b\u5b50\u8cc7\u6599\u593e")
    d.Add "scan.sum.ok", U("{0} \u500b\u6700\u65b0")
    d.Add "scan.sum.all.current", U("\u5168\u90e8\u6700\u65b0")
    d.Add "scan.sum.no.idw", U("\u627e\u4e0d\u5230\u5716\u9762")
    d.Add "scan.sum.missing", U("{0} \u500b\u7f3a\u5c11")
    d.Add "scan.sum.stale", U("{0} \u500b\u904e\u6642")
    d.Add "scan.sum.mismatch", U("{0} \u500b\u4e0d\u7b26")
    d.Add "scan.sum.to_create", U("{0} \u5efa\u7acb")
    d.Add "scan.sum.to_recreate", U("{0} \u91cd\u5efa")
    d.Add "scan.sum.force_suffix", U("(\u5f37\u5236)")
    d.Add "scan.validate.error", U("\u8cc7\u6599\u593e\u9a57\u8b49\u932f\u8aa4\u3002")
    d.Add "scan.sum.errors", U("{0} \u500b\u932f\u8aa4")
    d.Add "scan.no.folder", U("\u672a\u9078\u64c7\u8cc7\u6599\u593e\u3002")
    d.Add "scan.no.folder.exists", U("\u8cc7\u6599\u593e\u4e0d\u5b58\u5728\u3002")
    d.Add "scan.path.ok", U("\u5df2\u627e\u5230\u8cc7\u6599\u593e \u2014 \u6b63\u5728\u6383\u63cf...")
    d.Add "scan.path.too.long", U("\u8def\u5f91\u8d85\u904e {0} \u500b\u5b57\u5143\u7684\u9650\u5236\u3002")
    d.Add "progress.processing", U("\u6b63\u5728\u8655\u7406...")
    d.Add "progress.init", U("\u6b63\u5728\u521d\u59cb\u5316...")
    d.Add "progress.cancel", U("\u53d6\u6d88")
    d.Add "progress.cancelling", U("\u6b63\u5728\u53d6\u6d88...")
    d.Add "scan.models.detail", U("\u6709\u5716\u9762\uff1a{0} | \u7121\uff1a{1}")
    d.Add "status.idle", U("\u5c31\u7dd2")
    d.Add "status.scanning", U("\u6b63\u5728\u6383\u63cf\u8cc7\u6599\u593e...")
    d.Add "status.publishing", U("\u6b63\u5728\u767c\u4f48\uff1a{0} / {1}...")
    d.Add "pdf.err.open", U("\u7121\u6cd5\u958b\u555f IDW \u6a94\u6848\u3002")
    d.Add "pdf.err.pagematch", U("PDF \u6709 {0} \u9801\uff0c\u9810\u671f {1} \u9801\u3002")
    d.Add "pdf.err.noext", U("\u8def\u5f91\u7121\u6548 (\u7121\u526f\u6a94\u540d)\uff1a{0}")
    d.Add "pdf.skipped.exists", U("\u5df2\u7565\u904e (\u5b58\u5728)")
    d.Add "pdf.skipped.exists.pages", U("\u5df2\u7565\u904e (\u5b58\u5728) \u00b7 {0}/{1} \u9801")
    d.Add "pdf.err.addin", U("\u627e\u4e0d\u5230 PDF Translator \u589e\u76ca\u96c6\u3002")
    d.Add "pdf.err.nofile", U("PDF \u532f\u51fa\u672a\u5efa\u7acb\u6a94\u6848\u3002")
    d.Add "pdf.err.export", U("\u532f\u51fa\u5931\u6557\uff1a{0}")
    d.Add "pdf.err.bak.delete", U("\u7121\u6cd5\u522a\u9664\u820a\u5099\u4efd\uff1a{0}")
    d.Add "pdf.err.bak.move", U("\u7121\u6cd5\u5099\u4efd\u73fe\u6709 PDF\uff1a{0}")
    d.Add "pdf.err.write", U("\u7121\u6cd5\u5beb\u5165 PDF\uff1a{0}")
    d.Add "report.hdr.title", U("Batch PDF Publisher v{0} \u2014 \u532f\u51fa\u5831\u544a")
    d.Add "report.hdr.mode", U("\u6a21\u5f0f\uff1a{0}")
    d.Add "report.hdr.folder", U("\u8cc7\u6599\u593e\uff1a{0}")
    d.Add "report.hdr.subfolders", U("\u5b50\u8cc7\u6599\u593e")
    d.Add "report.hdr.force", U("\u5f37\u5236\u8986\u5beb")
    d.Add "report.hdr.generated", U("\u7522\u751f\u6642\u9593\uff1a{0}")
    d.Add "report.state.on", U("\u958b")
    d.Add "report.state.off", U("\u95dc")
    d.Add "report.hdr.pdf.section", U("PDF \u532f\u51fa")
    d.Add "report.hdr.pdf.detail", U("\u8a73\u7d30\u8cc7\u8a0a\uff1a")
    d.Add "report.created", U("\u5df2\u5efa\u7acb")
    d.Add "report.overwritten", U("\u5df2\u8986\u5beb")
    d.Add "report.skipped", U("\u5df2\u7565\u904e")
    d.Add "report.errors", U("\u932f\u8aa4")
    d.Add "report.time", U("\u6642\u9593")
    d.Add "report.empty", U("\u7a7a")
    d.Add "report.pages", U("\u9801")
    d.Add "report.no.views", U("\u7121\u8996\u5716")
    d.Add "report.tag.created", U("[\u5df2\u5efa\u7acb]")
    d.Add "report.tag.overwritten", U("[\u5df2\u8986\u5beb]")
    d.Add "report.tag.skipped", U("[\u5df2\u7565\u904e]")
    d.Add "report.tag.error", U("[\u932f\u8aa4]")
    d.Add "report.tag.hasdrawing", U("[\u6709\u5716\u9762]")
    d.Add "report.tag.nodrawing", U("[\u7121\u5716\u9762]")
    d.Add "report.hdr.audit.section", U("\u5716\u9762\u7a3d\u6838")
    d.Add "report.hdr.audit.detail", U("\u8a73\u7d30\u8cc7\u8a0a\uff1a")
    d.Add "report.audit.title", U("\u5716\u9762\u7a3d\u6838")
    d.Add "errdlg.retry.prompt", U("\u91cd\u8a66 = \u518d\u8a66 | \u5ffd\u7565 = \u7565\u904e | \u4e2d\u6b62 = \u505c\u6b62")
    d.Add "tip.run", U("\u958b\u59cb\u6279\u6b21 PDF \u532f\u51fa")
    d.Add "tip.theme", U("\u5207\u63db\u4e3b\u984c (Inventor / \u7cfb\u7d71 / \u6dfa\u8272 / \u6df1\u8272)")
    d.Add "tip.reset", U("\u5c07\u6240\u6709\u9078\u9805\u9084\u539f\u70ba\u9810\u8a2d\u503c")
    d.Add "tip.cancel", U("\u53d6\u6d88\u76ee\u524d\u64cd\u4f5c (ESC)")
    d.Add "tip.filter.nodrawing", U("\u50c5\u986f\u793a\u7121\u5716\u9762\u7684\u6a21\u578b")
    d.Add "tip.list", U("\u6309\u5169\u4e0b\u958b\u555f PDF\uff1b\u6309\u4e00\u4e0b\u67e5\u770b\u8a73\u7d30\u8cc7\u8a0a")
    d.Add "tip.subfolders", U("\u905e\u8ff4\u6383\u63cf\u5b50\u8cc7\u6599\u593e (\u7565\u904e OldVersions)")
    d.Add "tip.force", U("\u5373\u4f7f PDF \u6bd4 IDW \u65b0\u4e5f\u91cd\u65b0\u5efa\u7acb")
    d.Add "tip.audit", U("\u6aa2\u67e5\u6a21\u578b (.ipt/.iam) \u662f\u5426\u7f3a\u5c11\u5716\u9762")
    d.Add "tip.retry", U("\u91cd\u8a66\u6240\u6709\u932f\u8aa4\u7684\u6a94\u6848")
    d.Add "tip.export", U("\u5c07\u5831\u544a\u5132\u5b58\u5230\u6a94\u6848\u4e26\u5728\u8a18\u4e8b\u672c\u4e2d\u958b\u555f")
    d.Add "report.msg.close_blocked", U("\u8655\u7406\u671f\u9593\u7121\u6cd5\u95dc\u9589\u3002")
    d.Add "report.msg.open_failed", U("\u7121\u6cd5\u958b\u555f\uff1a\u627e\u4e0d\u5230\u6a94\u6848\u6216\u64cd\u4f5c\u88ab\u5c01\u9396\u3002")
    d.Add "report.msg.select_row", U("\u9078\u64c7\u5217\u4ee5\u67e5\u770b\u8a73\u7d30\u8cc7\u8a0a \u00b7 \u6309\u5169\u4e0b\u958b\u555f PDF")
    d.Add "report.msg.retry.none", U("\u91cd\u8a66\uff1a\u6c92\u6709\u8655\u65bc\u932f\u8aa4\u72c0\u614b\u7684\u6a94\u6848\u3002")
    d.Add "report.msg.retry.error", U("\u91cd\u8a66\u932f\u8aa4\uff1a{0}")
    d.Add "report.msg.retry.partial", U("\u91cd\u8a66\uff1a{0} \u500b\u6210\u529f\uff0c{1} \u500b\u4ecd\u932f\u8aa4\u3002")
    d.Add "report.msg.retry.all_fixed", U("\u91cd\u8a66\uff1a{0} \u500b\u6210\u529f\uff0c\u5168\u90e8\u5df2\u4fee\u5fa9\u3002")
    d.Add "report.msg.copied", U("\u5df2\u8907\u88fd\uff01")
    d.Add "report.msg.exported", U("\u5831\u544a\u5df2\u5728\u8a18\u4e8b\u672c\u4e2d\u958b\u555f\u3002")
    d.Add "report.msg.export_failed", U("\u7121\u6cd5\u5132\u5b58\u5831\u544a\u3002")
    d.Add "reset.msg.all_defaults", U("\u6240\u6709\u8a2d\u5b9a\u5df2\u9084\u539f\u70ba\u9810\u8a2d\u503c\u3002")
    d.Add "reset.lbl.registry", U("\u767b\u9304")
    d.Add "reset.lbl.sandbox", U("\u6c99\u7bb1")
    d.Add "reset.lbl.state", U("\u72c0\u614b")
    d.Add "reset.state.deleted", U("\u5df2\u522a\u9664")
    d.Add "reset.state.reinitialized", U("\u5df2\u91cd\u65b0\u521d\u59cb\u5316")
    d.Add "reset.msg.failed", U("\u91cd\u8a2d\u5931\u6557")
    d.Add "reset.msg.registry_complete", U("\u767b\u9304\u5df2\u91cd\u8a2d\u3002")
    d.Add "reset.msg.sandbox_complete", U("\u6c99\u7bb1\u5df2\u91cd\u8a2d")
    d.Add "scan.badge.tooltip", U("\u7e3d\u8a08\uff1a{0} | \u6700\u65b0\uff1a{1} | \u7f3a\u5c11\uff1a{2} | \u904e\u6642\uff1a{3} | \u4e0d\u7b26\uff1a{4} | \u932f\u8aa4\uff1a{5}")
    d.Add "report.msg.opening", U("\u6b63\u5728\u958b\u555f IDW \u6a94\u6848...")
    d.Add "report.msg.updating", U("\u6b63\u5728\u66f4\u65b0\u8996\u5716...")
    d.Add "report.msg.exporting", U("\u6b63\u5728\u532f\u51fa PDF...")
    d.Add "msg.no.write.access", U("\u7121\u6cd5\u5beb\u5165\u6c99\u7bb1\u8cc7\u6599\u593e\u3002\u8acb\u6aa2\u67e5\u6b0a\u9650\uff1a")
    d.Add "msg.pdf.temp_missing", U("\u532f\u51fa\u671f\u9593\u81e8\u6642 PDF \u6d88\u5931 (\u53ef\u80fd\u662f\u9632\u6bd2\u9694\u96e2\u6216\u6a94\u6848\u9396\u5b9a)\uff1a{0}")
    d.Add "msg.batch.abort.network", U("\u6279\u6b21\u4e2d\u6b62\uff1a\u9023\u7e8c {0} \u500b\u7db2\u8def\u932f\u8aa4\u3002\u8acb\u6aa2\u67e5\u9023\u7dda\u4e26\u91cd\u8a66\u3002")
    d.Add "msg.already.running", U("Batch PDF Publisher \u5df2\u5728\u57f7\u884c\u3002\u8acb\u95dc\u9589\u7b2c\u4e8c\u500b\u5de5\u4f5c\u968e\u6bb5\u4e26\u91cd\u8a66\u3002")
    d.Add "report.search.placeholder", U("\u641c\u5c0b\u8def\u5f91...")
    d.Add "report.search.label", U("\u641c\u5c0b\uff1a")
    d.Add "scan.collecting", U("\u6b63\u5728\u6536\u96c6\u6a94\u6848...")
    d.Add "scan.analysing", U("\u6b63\u5728\u5206\u6790\uff1a{0} / {1}...")
    d.Add "scan.auditing.start", U("\u6b63\u5728\u6aa2\u67e5\u6a21\u578b...")
    d.Add "scan.auditing", U("\u7a3d\u6838\uff1a{0} / {1}...")
    d.Add "scan.cancelled", U("\u5df2\u53d6\u6d88")
    d.Add "tip.search", U("\u6309\u8def\u5f91\u7be9\u9078\u5831\u544a\u5217")
    d.Add "msg.confirm.reset", U("\u5c07\u6240\u6709\u9078\u9805\u9084\u539f\u70ba\u9810\u8a2d\u503c\uff1f")
    d.Add "scan.disk.insufficient", U("\u78c1\u789f\u7a7a\u9593\u4e0d\u8db3\u3002\u53ef\u7528\uff1a{0}")
    d.Add "scan.title.no_folder", U("\u9078\u64c7\u8cc7\u6599\u593e")
    d.Add "summary.hdr.models_full", U("{0} \u500b\u6a21\u578b ({1} \u500b\u7d44\u4ef6\uff0c{2} \u500b\u96f6\u4ef6)")
    d.Add "summary.hdr.with_idw", U("{0} \u500b\u6709\u5716\u9762")
    d.Add "summary.hdr.no_idw", U("{0} \u500b\u7121\u5716\u9762")
    d.Add "errdlg.label.context", U("\u5167\u5bb9")
    d.Add "errdlg.label.time", U("\u6642\u9593")
    d.Add "errdlg.label.log", U("\u8a18\u9304")
    d.Add "errdlg.label.stack", U("\u5806\u758a")
    d.Add "pdf.err.prepare", U("\u6e96\u5099\u5931\u6557\uff1a{0}")
    d.Add "pdf.err.header", U("\u532f\u51fa\u7684\u81e8\u6642 PDF \u672a\u901a\u904e\u6a19\u982d\u9a57\u8b49")
    d.Add "pdf.err.unknown", U("\u672a\u77e5\u932f\u8aa4")
    d.Add "report.filter.all", U("\u5168\u90e8")
    d.Add "report.filter.err", U("\u932f\u8aa4")
    d.Add "report.filter.skip", U("\u5df2\u7565\u904e")
    d.Add "report.filter.ok", U("OK")
    d.Add "report.filter.noidw", U("\u7121 .idw")
    d.Add "scan.reset.sure", U("\u78ba\u5b9a\uff1f")
    d.Add "scan.reset.done", U("\u5b8c\u6210")
    d.Add "report.form.title", U("\u8655\u7406\u5831\u544a")
    d.Add "report.hdr.sub", U("\u5b50\u8cc7\u6599\u593e")
    d.Add "report.hdr.force_short", U("\u5f37\u5236")
    d.Add "report.hdr.audit", U("\u7a3d\u6838")
    d.Add "report.summary.full", U("{0} \u5df2\u5efa\u7acb \u00b7 {1} \u5df2\u8986\u5beb \u00b7 {2} \u5df2\u7565\u904e \u00b7 {3} \u500b\u932f\u8aa4 \u00b7 {4}")
    d.Add "report.skip.empty", U("{0} ({1} \u500b\u7a7a)")
    d.Add "scan.browse_title", U("\u9078\u64c7\u4f86\u6e90\u8cc7\u6599\u593e")
End Sub

' END: modLang v9.19.10
