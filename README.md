# Batch PDF Publisher

[![Version](https://img.shields.io/badge/version-9.25.32-blue?style=flat-square)](CHANGELOG.md)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)
[![Inventor](https://img.shields.io/badge/Inventor-2026%20%7C%202027-orange?style=flat-square)](https://www.autodesk.com/products/inventor/)
[![VBA](https://img.shields.io/badge/VBA-7.1-purple?style=flat-square)](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-for-applications)
[![Tests](https://img.shields.io/badge/tests-49%20passed-brightgreen?style=flat-square)](src/modTests.bas)
[![i18n](https://img.shields.io/badge/i18n-13%20languages-blueviolet?style=flat-square)](src/modLang.bas)
[![Last commit](https://img.shields.io/github/last-commit/iddj27510/BatchPDFPublisher?style=flat-square&label=last%20commit)](https://github.com/iddj27510/BatchPDFPublisher/commits/main)
[![CI](https://github.com/iddj27510/BatchPDFPublisher/actions/workflows/ci.yml/badge.svg)](https://github.com/iddj27510/BatchPDFPublisher/actions/workflows/ci.yml)

VBA add-in for Autodesk Inventor 2027. Bulk-converts `.idw` drawings
to PDF. Per-file report with filters and search. Optional drawing audit.

**Version:** 9.25.32
**Host:** Autodesk Inventor 2027 (also tested on 2026)
**Runtime:** VBA 7.1
**License:** MIT - see LICENSE

## Features

- Bulk IDW -> PDF export with skip / stale / mismatch detection
- Optional recursive folder scan (skips OldVersions)
- Optional drawing audit: .ipt / .iam with or without sibling IDW
- Optional QA validation of IDW views and PDF size
- 13 languages: en, de, fr, it, es, pt, ru, pl, cs, ja, ko, zh, zht
- 4 themes: Inventor / System / Light / Dark
- Structured logging with 5 levels, JSON mode, multi-archive rotation,
  and a separate 90-day audit trail
- Crash reports with call stack written to sandbox
- Single-instance lock (prevents two Inventor sessions from colliding)
- Network failure detection (5 consecutive -> abort with reason)
- Search box + status filters in report list
- Progress phases: Opening IDW / Updating views / Exporting PDF
- Scan options rendered as auto-sizing label toggles with Alt+
  accelerators (Alt+S / Alt+F / Alt+A)

## Recent improvements

### v9.25.32 - Scan options redesign + undo fix

- **Scan options are now single-label toggles** instead of
  CheckBoxes. `Include subfolders`, `Force recreate`, and
  `Audit drawings` are individual Labels with an Active/Inactive
  visual state (accent background when on, neutral when off).
  This removes three MSForms quirks that the CheckBox
  implementation had to work around (first-click after
  programmatic assignment, programmatic-Click side effect,
  duplicate Click from double-click).
- **Toggle labels auto-size to content on language switch.**
  `Alt+S` / `Alt+F` / `Alt+A` accelerators preserved.
- **No-flash form activation.** Registry read, caption setup,
  theme application, toggle layout, and toggle colouring moved
  from `UserForm_Activate` to `UserForm_Initialize`. First paint
  is the final state.
- **Action-oriented badge.** `lblSum` now says
  `27 drawings · 3 to create · 24 current` instead of
  `27 drawings · 3 missing`. Force ON moves all current files
  into the "to recreate" count and appends `(force)`.
- **Ctrl+Z in the folder path field now refreshes the summary.**
  MSForms consumes Ctrl+Z inside the Edit control before VBA
  sees it — the `Change` event does not fire. `RevalidatePathIfChanged`
  compares the text against the cached path and re-runs validation
  when it differs.
- **"No drawings found" is now shown in red**, not neutral grey —
  it is a blocking state, not a hint.
- **Audit toggle no longer flickers `lblSum`.** Toggling audit
  only updates `lblSum2`.
- **Test suite grew from 40 to 43 tests.** Three new tests:
  `Test_STS_Classify_ValidateStatus`,
  `Test_STATE_Init_ClearsRegCache`,
  `Test_PRJ_PutClipboardFallback_EnsuresSandbox`.
- **Known limitation:** pure keyboard Ctrl+Z with no subsequent
  mouse movement does not refresh the labels until the user moves
  the mouse. See `docs/KNOWN_LIMITATIONS.md`.

### v9.19.4 - Consistency and audit fixes

- **`STATE_ValidateEnum` substring bug fixed.** Prior code used
  `InStr(sAllowed, sVal)` which accepted substrings — `"IS"` and
  `"ISL"` passed against `"ISLD"`. Now uses delimited exact match.
- **Report file separators unified** to `" · "`.
- **`modFIT` skip list extracted** to `FIT_IsSkipped(sForm, sControl)`.

### v9.19.3 - Testing, refactoring, and consistency pass

- **Test suite expanded 15 -> 40** tests. New coverage: PDF header
  validation, temp path uniqueness, cache case-insensitivity,
  cache eviction at scale, relative path resolution, retry wrappers,
  QA size threshold, report filters, format boundaries, UNC path
  edge cases, and unicode decoder.
- **`modState.STATE_ValidateConfig` refactored** from monolithic
  block (CC ~12) to table-driven orchestrator plus four typed
  validators. Behavior preserved exactly (verified by test).
- **`modPath.SN_ResolveParentSegments` root-clamp fix.** Prior code
  dropped the drive prefix on paths like `C:\Foo\..\..\Bar`
  (returned `Bar` instead of `C:\Bar`). Now Windows semantics are
  applied: parent traversal is clamped at the root.
- **`modLog.LOG_OpenInternal` stale handle recovery.** A leftover
  file handle from a prior VBE session no longer blocks logging
  on startup. On error 55, `Close` is issued then the open retried.
- **Duplicate text writer removed.** `modReportExport` and
  `modState.STATE_BackupPreferences` now use the canonical
  `modSandbox.PRJ_WriteTextFile`. One code path for all text
  file writes.
- **Header standardization.** All modules and classes now use a
  uniform English header (Module/Purpose/Reference/Version).
- **`Uninstall_BatchPDFPublisher`** added — cleans registry and
  sandbox, prints the VBE component list for manual removal.
- **AUDIT log markers** for destructive operations
  (PDF_CommitReplace, PRJ_CleanupOrphans, STATE_ResetPreferences).

### v9.19.2 - Locale-aware dates

- **Date/time uses Windows Region settings.** `PRJ_FormatDate`
  and new `PRJ_FormatNow` now call `Format$(d, "Short Date")` and
  `Format$(d, "Short Time")` — output matches Explorer and every
  native app. Per-language hardcoded templates removed.
- **CJK encoding fix.** All `L_FillJA/KO/ZH/ZHT` use `U()` for
  non-ASCII characters; raw CJK literals no longer corrupt on
  CP1252 hosts.
- **Summary keys added** — `report.summary.full` and
  `report.skip.empty` existed in code but not in `modLang`.
  Summary line rendered as `??report.summary.full??` in every
  language. Fixed.

### v9.19.1 - UI text consolidation

- 7 keys added in all 13 languages: `chk.force`,
  `scan.reset.sure`, `scan.reset.done`, `report.form.title`,
  `report.hdr.sub`, `report.hdr.force_short`, `report.hdr.audit`.
- 4 dead keys removed.

### v9.19.0 - Audit summary + Inventor brand colors

- **Audit summary narrative.** `summary.hdr.models_full`
  replaces three separate keys. `with_idw` / `no_idw` reworded
  to "covered" / "missing".
- **Report form:** `lblSel` removed; row details live in
  `lstReport` columns. `lblTitle` becomes static.
- **Progress form:** no FIT on `lblFile`; designer font preserved.
- **Inventor brand palette** aligned with Autodesk "Color Register
  and Semantics" v4.1. All text colors pass WCAG AA in both themes.

### v9.18.8 - Keyboard hook disabled

- `frmReport`: `CKeyPreview` hook disabled. The classes remain
  in the project but are no longer instantiated on this form.

### v9.18.7 - Captions localized

- `PRJ_LocalizeCaptions` in `modUIHelpers`. Button captions now
  follow the `Lang` setting (previously the designer default
  stayed in place regardless of language).

### v9.18.6 - Accessibility documentation

- `docs/ACCESSIBILITY.md` added. Focus ring on `txtPath` and
  `txtSearch`. `PRJ_CheckAccessibility` advisory audit.

### v9.18.5 - Logging overhaul

- CRITICAL level, JSONL mode, multi-archive rotation (5 archives),
  separate 90-day audit file, runtime reload.

### v9.18.4 - Dependency lock file

- `deps.lock.json`, `Build-SBOM.ps1` v2.3.0 with mandatory hashes,
  `PRJ_CheckHostVersion`.

### v9.18.0 - Architecture consolidation + UX overhaul

Large folder scans (1000+ files) now show real-time progress,
support cancellation via ESC, and never block the UI.

| Scenario | Before | After |
| --- | --- | --- |
| 500 files, local SSD | 5s blocked | Instant form + live counter |
| 5,000 files, SMB share | 40s frozen | Responsive UI + ESC to cancel |
| 50,000 files, VPN | 15 min frozen | Live progress + cancel anytime |

## Requirements

- Autodesk Inventor 2027 (tested) or 2026 (tested)
- Autodesk Apprentice Server (recommended)
- Windows 7 or later (x86 or x64)
- VBA Enabler for Inventor (installed by Autodesk installer)

## Installation

### Option A: Import source files (development)

1. Open Inventor
2. Press `Alt+F11` (VBA Editor)
3. In VBE: `Tools -> Options -> General`, enable
   **Trust access to the VBA project object model**
4. In VBE: `File -> Import File...`
5. `Debug -> Compile VBAProject`
6. Run macro `Run_BatchPDFPublisher`

### Option B: Deploy as compiled .ivb (production)

1. Copy `BatchPDFPublisher.ivb` to
   `%APPDATA%\Autodesk\Inventor 2027\Addins\`
2. Place `BatchPDFPublisher.addin` manifest alongside
3. Restart Inventor
4. Run macro `Run_BatchPDFPublisher`

## Quick start

1. Run `Run_BatchPDFPublisher`
2. Select source folder
3. Optional toggles: Include subfolders, Force recreate, Audit drawings
4. Click **RUN**
5. Monitor progress; use **Cancel** to stop after current file
6. Review report. Use search box and filters
7. Use **Copy**, **Export**, **Retry**

## Report format

    Batch PDF Publisher v9.25.32 - Export Report
    Mode: PDF
    Folder: D:\Projects\PROJ-047\Drawings
    Subfolders: On  Force: Off
    Generated: 01.10.2026 14:32

    PDF EXPORT
    Created: 45  Overwritten: 0  Skipped: 3  Errors: 2  Time: 43.21 s

    Details:
    [Created] - 245 KB - Bracket\Bracket.idw
    [Error] - - - Corrupt\Hairdryer.idw
    ...

Per-file status prefixes:

- `[Created]` - new PDF produced
- `[Overwritten]` - existing PDF replaced
- `[Skipped]` - PDF already current
- `[Error]` - export failed
- `[Has Drawing]` / `[No Drawing]` - audit mode only

## Configuration

Registry: `HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs`

| Key | Default | Description |
| --- | --- | --- |
| `Theme` | `I` | I=Inventor, S=System, L=Light, D=Dark |
| `Path` | "" | Last used source folder |
| `Sub` | True | Include subfolders |
| `Force` | False | Force recreate |
| `Audit` | False | Audit mode on |
| `LogLevel` | 1 | 0=DEBUG, 1=INFO, 2=WARN, 3=ERROR, 4=CRITICAL |
| `LogJson` | False | True = JSONL output; False = plain text |
| `Lang` | "" | Empty = auto-detect; else en/de/fr/... |
| `SchemaVersion` | 2 | Registry schema version (STATE_MigrateSchema) |

## Sandbox and logging

Sandbox: `%TEMP%\VBADR\`

Contains:

- `BatchPDF_Log_<PID>.txt` - per-process log, rotates at 1 MB
- `BatchPDF_Log_prev1_<PID>.txt` ... `prev5_<PID>.txt` - archives
- `BatchPDF_Audit_<PID>.txt` - audit trail (90-day retention)
- `_tmp_*.pdf` - temporary export files
- `crash_<timestamp>.txt` - crash reports
- `_instance.lock` - single-instance lock

Orphan cleanup runs at startup: main log files and temp files older
than 24h, `.bak` files older than 7 days, audit files older than 90
days.

### Runtime log reconfiguration

    PRJ_ReloadLogLevel

Re-reads `LogLevel` and `LogJson` from registry without restarting
Inventor.

## Testing

From VBA Immediate Window (`Ctrl+G`):

    RunAllTests

Expected output:

    TOTAL: 43 / 43 passed, 0 failed

Test inventory (43 tests):

| # | Test | Area |
|---|---|---|
| 1 | SN_SanitizePath | Path sanitization |
| 2 | PDF_SafeReplace | File I/O |
| 3 | RP_StatusTag | Report status |
| 4 | CACHE_RoundTrip | Cache |
| 5 | L_FallbackChain | Localization |
| 6 | STS_Classify | Status pipeline |
| 7 | PRJ_FormatElapsed | Formatting |
| 8 | PRJ_FormatFileSize | Formatting |
| 9 | PRJ_FormatDate | Formatting |
| 10 | RP_MatchesFilter | Report filters |
| 11 | STATE_AssertGlobals | Global state |
| 12 | ERR_Classify | Error classification |
| 13 | PRJ_PathCombine | Path helpers |
| 14 | RP_BuildSummaryText | Localized summary |
| 15 | STATE_ValidateConfig | Config validation |
| 16 | PDF_ValidateHeader | PDF magic bytes |
| 17 | PDF_BuildTempPath | Temp path structure |
| 18 | CACHE_CaseInsensitive | Cache dictionary |
| 19 | PRJ_FormatElapsed_Edge | Timing boundary |
| 20 | PRJ_FormatDate_Edge | Date FAT epoch |
| 21 | SN_SanitizePath_UNC | UNC edge cases |
| 22 | ERR_IsRecoverable_All | Error recovery |
| 23 | ERR_DescribeClass_All | Error labels |
| 24 | L_Placeholders | L2/L3/LN |
| 25 | U_UnicodeDecoder | \uXXXX decoder |
| 26 | LANG_NormalizeCode | Lang code normalization |
| 27 | RP_CountFiles | Synthetic counts |
| 28 | RP_PickSummaryColor | Summary color |
| 29 | RP_PickStatusBold | Bold flag |
| 30 | RP_StatusColor | Status colors |
| 31 | CACHE_EvictIfNeeded | Cache eviction |
| 32 | PRJ_GetRelPath | Relative path |
| 33 | RETRY_DeleteFile_MoveFile | Retry wrappers |
| 34 | QA_AnalyzePDF_TooSmall | QA threshold |
| 35 | RP_FindStat | Report lookup |
| 36 | RP_CountErrorRows | Error counting |
| 37 | FormatFileSize_Boundary | Unit transitions |
| 38 | STS_Classify_Edge | Status edge cases |
| 39 | LANG_SupportedCodes | Language list |
| 40 | FormatElapsed_Boundary | Time boundaries |
| 41 | STS_Classify_ValidateStatus | Status pipeline (Missing/Stale/Mismatch/OK) |
| 42 | STATE_Init_ClearsRegCache | Registry cache reset |
| 43 | PRJ_PutClipboardFallback_EnsuresSandbox | Clipboard fallback regression guard |

Also available:

    PRJ_SelfTest                 ' Environment + accessibility audit
    PRJ_CheckHostVersion         ' Returns True if Inventor is 2026/2027
    PRJ_CheckAccessibility       ' Standalone accessibility audit
    Uninstall_BatchPDFPublisher  ' Clean external state, list VBE components

## Architecture

The project is organized into **39 deliverable modules** across
**6 layers**:

- **Core** - modConstants, modState, modEntry
- **Scan & PDF** - modBatch, modPDFExport, modPDFFileIO,
  modPDFProgress, modInspect, modValidate, modStatus, modQA, modRetry
- **Report** - modReportBuild, modReportExport, modReportFilter,
  modSummary
- **Infrastructure** - modLog, modError, modSandbox, modCOM,
  modCache, modFileInfo, modPath, modTime, modWinAPI, modSelfTest
- **Localization** - modLang (13 languages), modHelp (13 languages)
- **UI** - modTheme, modPalette, modHover, modFIT, modUIHelpers,
  modClipboard, modBrowse, modShellApp, CHoverBtn (class),
  CKeyPreview + CKeyPreviewHook + 5 hook classes (classes),
  frmProgress, frmReport, frmScan (UserForms)

Dev-only modules (never in `src/` or Git): `a_modDiag`, `a_modExport`.

See `docs/ARCHITECTURE.md` for full details.

## Documentation

- `docs/ARCHITECTURE.md` - module map, invariants, concurrency,
  testability, maintainability, dependency management, logging
- `docs/ACCESSIBILITY.md` - WCAG 2.2 AA mapping, EN 301 549,
  screen reader behaviour, accepted limitations
- `docs/KNOWN_LIMITATIONS.md` - documented limits and workarounds
- `docs/ROLLBACK.md` - rollback procedure
- `docs/INVENTOR_API_GOTCHAS.md` - Inventor API quirks
- `docs/INVENTOR_2027_NOTES.md` - Inventor 2027 API notes
- `COMPLIANCE.md` - WCAG status, EU CRA, OWASP mapping
- `SECURITY.md` - vulnerability reporting, CSIRT SLA
- `CHANGELOG.md` - version history
- `ROADMAP.md` - planned features
- `TECH_DEBT.md` - tracked technical debt

## Test fixtures

`T01_HappyPath`, `T04_Recursive`, `T05_Collision`, `T07_ErrorHandling`,
`T08_EmptyFolder`, `T10_MultiSheet`, `T11_DWG` contain real Inventor
files (~107 MB total). Not tracked in Git.

## Accessibility

This add-in targets WCAG 2.2 AA within MSForms platform constraints.
Full conformance detail in `docs/ACCESSIBILITY.md`.

What passes:

- Colour contrast in both themes (light and dark)
- Keyboard navigation: all actions reachable without a mouse
- Logical tab order (left-to-right, top-to-bottom)
- ESC handler on every form
- Focus ring on txtPath and txtSearch (BorderColor changes on focus)
- Access keys for primary actions (Run, Browse, Theme, Reset, Cancel,
  Subfolders, Force, Audit)
- ControlTipText on all interactive controls
- Error messages related to the relevant input field (txtPath)
- Reset confirmation dialog for destructive actions
- High contrast mode detection

What does not pass (MSForms platform limitations):

- Focus indicator on Label-based buttons (platform constraint)
- Screen reader `AccessibleName` / `AccessibleDescription`
- DPI scaling above 150%
- Live region announcements for dynamic status updates
- Scan option toggles have no announced state (Labels, not CheckBoxes)

Full details in `docs/ACCESSIBILITY.md` and `COMPLIANCE.md`.

## Author

**
Jugoslav Ilić
**  
Email: 
engilic@gmail.com

Batch PDF Publisher is designed, developed, and maintained by 
Jugoslav Ilić
.

## License

MIT - see LICENSE file for details.