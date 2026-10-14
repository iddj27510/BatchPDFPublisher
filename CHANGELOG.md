# Changelog

[9.25.42] - 2026-10-14
### Fixed
- AutoFitListBoxHeight: IntegralHeight snap fix eliminates partial-row scroll bug.
  When filter returns exactly viewport rows (e.g. OK filter with 26 rows),
  scroll range=0 no longer occurs.

[9.25.41] - 2026-10-14
### Fixed
- MouseScroll: ProcessMouseData now accepts Static/Button/Edit class names.
  Scroll wheel works when hovering over Label/Button controls.

[9.25.40] - 2026-10-14
### Changed
- Status localization via RP_LocalizedStatusBase in frmReport.

[9.25.39] - 2026-10-14
### Fixed
- Silent guard when g_colFiles is Nothing in frmReport.
- APP_TITLE integration in frmReport.

[9.25.38] - 2026-10-14
### Fixed
- OK filter scroll fix attempt (DoEvents/SetFocus in RP_SetFilter).

[9.25.37] - 2026-10-14
### Fixed
- AutoFitListBoxHeight: IntegralHeight ordering fix.
  Sets IntegralHeight=False BEFORE assigning lst.Height to prevent
  scrollbar hiding when list grows after filter switch.

[9.25.36] - 2026-10-14
### Added
- Mouse wheel scroll on frmProgress and frmReport (EnableMouseScroll/DisableMouseScroll).
- Status localization via RP_LocalizedStatusBase.

[9.25.35] - 2026-10-14
### Changed
- APP_TITLE constant integration across all forms and error handlers.
  Product name never translated, identical across all locales.

[9.25.34] - 2026-10-14
### Fixed
- FR/PT diacritics in reset.* keys (U() helper for non-ASCII).

[9.25.33] - 2026-10-14
### Changed
- Skipped status shortened: "Skipped (exists) - X/Y pages" → "Skipped - X/Y".
- CJK forms use middle-dot separator to match page_format keys.


[9.25.42] - 2026-10-14
### Added
- modReportList module (v1.0.0) - Label-based grid replacing ListBox.
- modWheelHook module - Native subclassing for mouse wheel scroll.
- Dynamic column width measurement using GDI (FONT_TextWidthPt).
- Thin scrollbar without arrows (Windows 11 / VS Code style).
- RL_BuildCells as single source of truth for display strings.

### Changed
- rmReport migrated from ListBox to Label grid (Opcija B).
- Removed MouseScroll.bas, MouseOverControl.cls, modListBoxFit.bas (~3000 LOC).
- Header heights measured once via FONT_CellPt.
- Data column widths measured from actual g_colFiles content.
- Final column widths = MAX(header_min, data_min).
- Path column elastic (takes remaining space).
- Frame height adjusted to exact row count (no partial rows).

### Fixed
- IntegralHeight partial-row bug eliminated by construction.
- MouseScroll hook crashes eliminated (no hook needed).
- Column width jumping on filter change eliminated (widths fixed after first load).

### Removed
- MouseScroll.bas (~1100 LOC).
- MouseOverControl.cls (~200 LOC).
- modListBoxFit.bas (~1700 LOC).
- lstReport ListBox control from rmReport designer.

### Notes
- lblListFrame serves as both visual border and input surface (D37).
- lblScrollTrack and lblScrollThumb added to designer for scrollbar.
- Data labels created programmatically on form (not inside frame).
- Z-order: frame → track → thumb → data labels.


All notable changes to Batch PDF Publisher.

Format based on [Keep a Changelog](https://keepachangelog.com/).

## [9.25.42] - 2026-10-13

### Fixed

- ZH translation: `report.list.path.empty` was the Korean string. Replaced with correct Simplified Chinese.
- `RP_ResetPathHint` called the non-existent key `report.list.path.hint`. Corrected to `report.list.path.empty`.
- `LANG_KeyForControl` for `lblListPath` used the same wrong key.

## [9.25.31] - 2026-10-13

### Added

- `AutoFitTruncateForCellKeepFilename`. Keeps full filename and extension; truncates folder portion.

## [9.25.30] - 2026-10-13

### Added

- `AutoFitTruncateForCellKeepExt`. Truncates while preserving file extension.

## [9.25.29] - 2026-10-13

### Changed

- iLogic suppression during batch. Prevents Inventor crash on E_FAIL from iLogic OnCloseDocument.

### Removed

- MouseScroll hook permanently removed (startup crash fix).

## [9.25.28] - 2026-10-13

### Added

- Master log switch `LOG_ENABLED`. Gates entire logging subsystem.

### Fixed

- Copy flicker in `frmReport.CopyClick`.

## [9.25.27] - 2026-10-13

### Changed

- `RP_ResetPathHint` + `EnableMouseScroll` position cleanup.

## [9.25.26] - 2026-10-13

### Fixed

- Height drift in ListBox auto-fit (locked after first fit).

## [9.25.25] - 2026-10-13

### Changed

- Payload-only columns hidden in `frmReport.lstReport`.

## [9.25.24] - 2026-10-13

### Added

- `AutoFitProbeRowHeight`. One-shot true FM20 row-height measurement.

## [9.25.23] - 2026-10-13

### Changed

- Phase 2/3 rewrite of listbox fit engine.

## [9.25.22] - 2026-10-13

### Added

- `AutoFitTruncateForCell`. Ellipsis-truncate a string to fit a column width.

### Fixed

- `CELL_PADDING_PX` raised 6 to 12. Copy flash fixes.

## [9.25.21] - 2026-10-13

### Changed

- OPTION B: px-grid geometry for ListBox columns.

## [9.25.20] - 2026-10-13

### Added

- `modListBoxFit` integration. Universal ListBox geometry.
- `modFont`. Central font policy + GDI precise text measurement.

### Changed

- `frmProgress` bold at completion only. Initial color = CLR_ACCENT, not CLR_ERROR. No red flash on open.
- `modFIT FIT_FitAllLabels` fits only labels ending in Txt with paired background.
- Badge fix: `SC_ApplyDrawingBadge` targets `lblSum`.

## [9.25.19] - 2026-10-13

### Changed

- Magic Numbers Cleanup: PRJ_AlignTxt uses constants.
- Dark-mode detection consolidated in modPalette.
- TPdfTiming.Total renamed to dTotal.
- VBA Optional parameter fix (literal 5).

## [9.25.7] - 2026-10-13

### Added

- Progress dialog layout keys (3 per language, 13 languages).

### Fixed

- Startup crash fix: single Load frmReport.

## [9.25.6] - 2026-10-13

### Added

- `CLR_NEUTRAL` = &H767676. Theme-invariant neutral for PdfColor B&W toggle.

### Changed

- frmScan.lblTitle is static (SC_SetFormTitle removed).

### Fixed

- Encoding + translation (ZHT, DE).
- MouseScroll idempotent EnableMouseScroll.

## [9.25.5] - 2026-10-13

### Changed

- Debug cleanup in hook classes.
- scan.title.no_folder added to critical keys test.

## [9.25.3] - 2026-10-13

### Changed

- Single-color toggle policy.

## [9.25.0] - 2026-10-13

### Added

- PDF options: DPI (300/600/1200) and Color/B&W toggles. 4 new lang keys.

## [9.22.5] - 2026-10-02

### Changed

- FormatFileSize alignment (one decimal place).

## [9.22.4] - 2026-10-02

### Changed

- Clean output (removed decorative separators).

## [9.22.2] - 2026-10-02

### Changed

- Version alignment.

## [9.22.1] - 2026-10-02

### Changed

- Unified one-decimal size format.
- Elapsed ticker rides scan/audit pump.

## [9.22.0] - 2026-10-02

### Changed

- Version alignment.

## [9.20.1] - 2026-10-01

### Fixed

- Audit classification unified (PRJ_HasDrawing for all three audit paths).

## [9.20.0] - 2026-10-01

### Added

- Audit rows in g_colFiles.
- Localized page/pages in PDF_BuildSuccessStatus.

## [9.19.9] - 2026-10-04

### Changed

- **frmScan scan options converted from CheckBox to single-label
  toggles.** The three options (`Include subfolders`, `Force
  recreate`, `Audit drawings`) were MSForms CheckBox controls.
  They are now individual Labels with an Active/Inactive visual
  state driven by `ApplyToggleColors`. This eliminates three
  MSForms quirks that were stacking on top of each other:
  (1) first-click quirk after programmatic `.Value` assignment,
  (2) `Click` fired as a side effect of programmatic assignment,
  (3) duplicate `Click` from Windows double-click handling. The
  workaround stack (`m_bInitializing` guard, debounce gate, flip +
  force-sync) is gone — Labels have no `.Value` and no auto-toggle.

- **Toggle labels auto-size to content on language switch.**
  `LayoutToggleLabels` measures each label's natural width via
  `AutoSize`, sets `Width = Ceiling(natural) + TOGGLE_WIDTH_PAD`
  (6pt) and `Height = designer height`. Position is chained:
  `Left = previous.Left + previous.Width + TOGGLE_GAP` (6pt).
  Re-runs on Activate, language switch, theme switch, and reset.

- **No-flash activation.** Registry read, caption setup, theme
  application, toggle layout and toggle colouring moved from
  `UserForm_Activate` to `UserForm_Initialize`. The first paint
  is already the final state — no designer-default → theme →
  accent transition is visible on form open.

- **Accelerators for scan option toggles:** `Alt+S` (Subfolders),
  `Alt+F` (Force), `Alt+A` (Audit). Labels do not receive Tab
  focus in MSForms; accelerators are the keyboard affordance.

- **`SC_BuildBadgeText` action-oriented rewrite.** The visible
  `lblSum` caption now describes what the batch will do rather
  than what is wrong:
  - Old: `27 drawings · 3 missing`
  - New: `27 drawings · 3 to create · 24 current`
  When `Force` is on, all current files move into the "to
  recreate" count and the caption ends with `(force)`. Errors
  stay visible in both modes. The diagnostic tooltip keeps the
  missing / stale / mismatch vocabulary. Three new lang keys:
  `scan.sum.to_create`, `scan.sum.to_recreate`,
  `scan.sum.force_suffix` — added to all 13 languages.

- **`OptForceClick` refreshes the badge immediately.**
  `SC_ApplyDrawingBadge` is called with the cached `g_tLastBadge`
  counters, so toggling Force updates the caption without a
  rescan and without flicker.

- **`frmScan.UserForm_Activate` runs `ApplyToggleColors` before
  `PRJ_UpdateThemeLabel`.** `PRJ_UpdateThemeLabel` internally
  triggers `FIT_FitLabel` → `FIT_MeasureText` → `DoEvents`, which
  forces the first paint. Colouring the toggles after that
  produced a visible designer-default → theme flash on open.

- **`OptAuditClick` no longer runs `SC_Validate`.** Toggling the
  audit option does not change the drawing set or their status.
  Only `SC_UpdateSum2` runs (populates or clears `lblSum2`). The
  visible `lblSum` badge no longer flickers when the user toggles
  audit.

### Fixed

- **Ctrl+Z in `txtPath` did not refresh `lblSum` / `lblSum2`.**
  MSForms consumes Ctrl+Z inside the Windows Edit control before
  VBA sees it — the `Change` event does not fire. The validation
  label stayed stale after the user restored deleted text.
  Workaround: `RevalidatePathIfChanged` compares the current text
  against `g_sLastFolder` and re-runs `SC_QuickCheck` +
  `SC_Validate` + `SC_UpdateSum2` when they differ. Called from
  `UserForm_MouseMove`.

- **`lblSum2` was stale after Undo.** The `RevalidatePathIfChanged`
  helper initially skipped `SC_UpdateSum2`; the audit summary
  label did not refresh. Now both `lblSum` and `lblSum2` update
  in the same pass.

- **`AutoSize = False` was assigned after `.Width` / `.Height`
  in `LayoutToggleLabels`.** MSForms ignores those assignments
  while `AutoSize` is True, so the +6pt padding had no effect.
  Property order corrected: `AutoSize = False` first, then
  `Width` / `Height`.

- **"No drawings found" colour changed from `CLR_TEXT_SEC` to
  `CLR_ERROR`.** Empty folder is a blocking state (batch cannot
  start), not a neutral one. Now renders red in both themes.

- **Host version guard used the wrong major-version numbers
  (BUG-02).** Prior code compared `SoftwareVersion.Major` against
  26 and 27 believing those were Inventor 2026 and 2027. Inventor
  uses an offset of 1996: 2026 is Major 30, 2027 is Major 31.
  Added `HOST_IsSupportedMajor(nMajor)` as the single source of
  truth. `PRJ_CheckHostVersion` (modEntry) and `PRJ_SelfTest`
  (modSelfTest) now delegate to it.

- **`frmProgress` and `frmReport` `UserForm_MouseMove` no longer
  call `FW_ApplyTheme`.** The call was triggering a visible
  repaint on every mouse move (flicker). `FW_ResetHover` remains.

- **`STATE_ValidateEnum` substring bug.** Prior implementation
  used `InStr(sAllowed, sVal) > 0`, which accepted any substring
  of the allow-list string. `"IS"` and `"ISL"` both passed
  against `sAllowed = "ISLD"` even though neither is a valid
  theme code, and the invalid value stayed in the registry.
  `FW_IsDark` then fell through to `Case Else` and silently
  returned the Inventor default instead of the configured theme.
  Now uses delimited exact match: `InStr("|" & sAllowed & "|",
  "|" & sVal & "|")`. Empty values always fail.

- **Report file separator inconsistency.** `RP_BuildPdfDetail`
  and `RP_BuildAuditDetail` in modReportBuild mixed `" · "` with
  `vbTab` between fields. Report files now use `" · "`
  throughout, matching the UI and the rest of the report body.

### Added

- **`ST_MISSING`, `ST_STALE`, `ST_MISMATCH`, `ST_OK` status enum
  values** in modStatus. Consumed by `SC_ValidateDrawings`
  (modValidate) for the validate/check-status pipeline. Previously
  `SC_ValidateDrawings` used a `Select Case True` block with mixed
  string equality and `Left$` prefix matching. Now unified through
  `STS_Classify`.

- **Audit count cache** in modValidate. `SC_ValidateAudit` stores
  the per-extension counts and the two totals (with/without
  drawing). `SC_UpdateSum2` checks the cache first and skips the
  four-pass `PRJ_CollectByExt` scan if the folder and subfolder
  flag match.

- **Three new tests** (test suite 40 -> 43):
  - `Test_STS_Classify_ValidateStatus` — verifies the four new
    validate statuses (Missing / Stale / Mismatch / OK) and
    rejects extended strings like `"MissingFile"`.
  - `Test_STATE_Init_ClearsRegCache` — verifies that `STATE_Init`
    clears the registry cache (`REG_ClearCache` at the top of the
    procedure). Without it, values cached in one session would
    survive a reset.
  - `Test_PRJ_PutClipboardFallback_EnsuresSandbox` — verifies that
    the fallback clipboard writer creates the sandbox itself when
    it has been wiped. Regression guard for the modClipboard
    `PRJ_EnsureSandbox` addition.

### Removed

- **`modScanPoll` (Win32 SetTimer poller).** Attempted to solve
  the Ctrl+Z problem with a background timer. Abandoned:
  Inventor VBA does not support the `AddressOf` operator
  (compile error 35010), which is required for the Win32
  callback. Removed before commit.

- **`STATE_ValidateEnum` substring matching.** Replaced with
  delimited exact match. See Fixed section.

### Notes

- **Versions 9.19.5 through 9.19.8 were internal iteration
  stamps** during this development cycle. They were never
  published as standalone releases. The cumulative changes are
  rolled into this 9.19.9 entry. `PRJ_VERSION` remained at
  9.19.4 in the shipped `modConstants.bas` until the final bump
  to 9.19.9.

- **Known limitation:** pure keyboard Ctrl+Z with no subsequent
  mouse movement does not refresh the labels until the user moves
  the mouse. In practice the user always moves the mouse or
  presses Enter; both actions refresh the labels. See
  `docs/KNOWN_LIMITATIONS.md` section "Ctrl+Z in txtPath" for the
  full analysis and the rejected alternatives.

### Test suite

- 43 tests, ~118 assertions, all passing. Three new tests added
  this cycle (see Added section). The remaining 40 tests and
  their assertions are unchanged from v9.19.4. Manual
  verification of UI changes: T01_HappyPath with all three
  toggles, plus Ctrl+Z behaviour in txtPath.

## [9.19.4] - 2026-10-03

### Fixed

- **`STATE_ValidateEnum` substring bug.** Prior implementation
  used `InStr(sAllowed, sVal) > 0`, which accepted any substring
  of the allow-list string. `"IS"` and `"ISL"` both passed
  against `sAllowed = "ISLD"` even though neither is a valid
  theme code, and the invalid value stayed in the registry.
  `FW_IsDark` then fell through to `Case Else` and silently
  returned the Inventor default instead of the configured theme.
  Now uses delimited exact match: `InStr("|" & sAllowed & "|",
  "|" & sVal & "|")`. Empty values always fail.
- **Report file separator inconsistency.** `RP_BuildPdfDetail`
  and `RP_BuildAuditDetail` in modReportBuild mixed `" · "` with
  `vbTab` between fields. Report files now use `" · "`
  throughout, matching the UI and the rest of the report body.

### Changed

- **`modFIT.FIT_FitAllLabels` skip list extracted.** The
  hardcoded condition `If Not (sFrm = "frmScan" And sName =
  "lblBrowseTxt")` buried in the middle of the control loop is
  now a private `FIT_IsSkipped(sForm, sControl)` helper. Adding
  a new skip entry is a single `Case` line; the reason for each
  entry is documented in the helper's comment block.
- **Form header notes clarified.** `frmScan`, `frmProgress`, and
  `frmReport` headers previously said "No line-continuation
  characters in this module", which conflicted with the
  `UserForm_MouseMove` signature (which uses `_`). The note now
  reads "No line-continuation characters in procedure bodies.
  Procedure signatures are excepted per Doc 01 / Doc 06." This
  matches the actual rule in the project style guide.
- **`frmScan.ResetClick` audit comment added.** A one-line comment
  now notes that `STATE_ResetPreferences` writes the AUDIT record
  itself, and that the call site must not duplicate it. Prevents
  a future maintainer from adding a second audit log entry.
- **CHANGELOG test count correction.** The v9.19.3 entry said
  "15 new tests (test suite 25 -> 40)" but the actual delta was
  25 (15 -> 40), split across two batches (Batch 4 added 15,
  Batch 4.5 added 10). The v9.19.3 entry has been corrected to
  list both batches and the full test inventory.

### Test suite

- 40 tests, 112 assertions (was 110), all passing.
- `Test_STATE_ValidateConfig` extended with two regression
  assertions for the `STATE_ValidateEnum` fix: `"IS"` and `"ISL"`
  must be rejected as theme codes and reset to `"I"`.

## [9.19.3] - 2026-10-03

### Added

- **25 new tests** (test suite 15 -> 40), added across two
  batches in this release cycle.

  Batch 4 (15 tests): `Test_PDF_ValidateHeader`,
  `Test_PDF_BuildTempPath`, `Test_CACHE_CaseInsensitive`,
  `Test_PRJ_FormatElapsed_Edge`, `Test_PRJ_FormatDate_Edge`,
  `Test_SN_SanitizePath_UNC`, `Test_ERR_IsRecoverable_All`,
  `Test_ERR_DescribeClass_All`, `Test_L_Placeholders`,
  `Test_U_UnicodeDecoder`, `Test_LANG_NormalizeCode`,
  `Test_RP_CountFiles`, `Test_RP_PickSummaryColor`,
  `Test_RP_PickStatusBold`, `Test_RP_StatusColor`.

  Batch 4.5 (10 tests): `Test_CACHE_EvictIfNeeded`,
  `Test_PRJ_GetRelPath`, `Test_RETRY_DeleteFile_MoveFile`,
  `Test_QA_AnalyzePDF_TooSmall`, `Test_RP_FindStat`,
  `Test_RP_CountErrorRows`, `Test_FormatFileSize_Boundary`,
  `Test_STS_Classify_Edge`, `Test_LANG_SupportedCodes`,
  `Test_FormatElapsed_Boundary`.

  Plus one fix in `Test_L_Placeholders`: L3 requires three args
  (key + two placeholders); the previous test called it with two.

- **`Uninstall_BatchPDFPublisher`** in modEntry. Cleans registry
  tree and sandbox folder, prints the list of VBE components
  for manual removal (VBE cannot safely remove the currently
  executing module).
- **AUDIT log markers** for destructive operations:
  `PDF_CommitReplace`, `PRJ_CleanupOrphans`,
  `STATE_ResetPreferences`.
- **UI Contract** integrated into Doc 01 as a standalone section.
- **modTests SIZE NOTE** header — module exceeds 800 LOC as a
  documented exception.
- **`PRJ_LogReloadJson`** — runtime JSON mode reload without
  restarting Inventor.

### Changed

- **`modState.STATE_ValidateConfig` refactored** from a monolithic
  block (CC ~12) to a flat orchestrator plus four typed validators
  (`STATE_ValidateEnum`, `STATE_ValidateBool`,
  `STATE_ValidateIntRange`, `STATE_ValidateLang`). Behavior
  preserved exactly — verified by `Test_STATE_ValidateConfig`.
- **`modReportExport` and `modState.STATE_BackupPreferences`**
  now use the canonical `modSandbox.PRJ_WriteTextFile`. Three
  duplicate private procedures removed.
- **Header standardization** — all modules and classes now use
  the English `Module/Purpose/Reference/Version` style. Style A
  (NAZIV/TIP/OPIS) headers converted.
- **`modStatus` magic numbers** replaced with `Len()` calls via
  `Const` declarations.
- **PRJ_VERSION bumped 9.19.2 -> 9.19.3.**

### Fixed

- **`modPath.SN_ResolveParentSegments` root-clamp bug.** Prior
  code dropped the drive prefix on paths like `C:\Foo\..\..\Bar`
  (returned `Bar` instead of `C:\Bar`). Windows semantics now
  applied: parent traversal is clamped at root.
- **`modLog.LOG_OpenInternal` stale handle recovery.** A leftover
  file handle from a prior VBE session no longer blocks logging.
  On error 55, `Close` releases all process file handles, then
  the open is retried once.
- **Missing footers** in modQA and modShellApp.
- **Corrupted characters** in modHelp header (`ï¿½4.3` → `§4.3`).
- **`Test_L_Placeholders`** — `L3` requires three arguments
  (key + two placeholders); the test was calling with two.

### Test suite

- 40 tests, all passing (`RunAllTests`).

## [9.19.2] - 2026-10-02

### Added

- `PRJ_FormatNow()` in modFileInfo — current wall-clock time in
  Windows locale format.
- `report.summary.full` and `report.skip.empty` keys in all 13
  languages (previously referenced in code but not defined;
  summary line rendered as `??report.summary.full??` in every
  language before this fix).

### Changed

- **Date/time now uses Windows Region settings.** `PRJ_FormatDate`
  calls `Format$(d, "Short Date")` and `Format$(d, "Short Time")`.
  Per-language hardcoded templates (`report.date.format`,
  `report.date.timeonly`) removed from all 13 languages.
- **`modReportBuild.RP_BuildHeader`** uses `PRJ_FormatNow`.
- **`modReportBuild.RP_BuildAuditBody`** rewritten to use the same
  unified summary keys as `modSummary.PRJ_GetAuditSummary`.
- **CJK encoding fix** — `L_FillJA/KO/ZH/ZHT` use `U()` exclusively
  for non-ASCII characters. Raw CJK literals no longer corrupt on
  CP1252 hosts.
- PRJ_VERSION bumped 9.19.1 -> 9.19.2.

### Test suite

- 25 tests, all passing.

## [9.19.1] - 2026-10-02

### Added

- 7 new UI text keys in all 13 languages: `chk.force`,
  `scan.reset.sure`, `scan.reset.done`, `report.form.title`,
  `report.hdr.sub`, `report.hdr.force_short`, `report.hdr.audit`.

### Removed

- 4 dead lang keys: `scan.sum.models`, `scan.sum.all.covered`,
  `scan.audit.sum.with`, `scan.audit.sum.without`.

### Changed

- PRJ_VERSION bumped 9.19.0 -> 9.19.1.

### Test suite

- 25 tests, all passing.

## [9.19.0] - 2026-10-02

### Added

- **Inventor brand palette** aligned with Autodesk "Color Register
  and Semantics" v4.1. Dark and Light themes documented with WCAG
  AA contrast ratios. `CLR_ACCENT` = Inventor blue `#2A8FD4`;
  `CLR_RUN_BG` darker for AA compliance with white text.
- **Tag-based theme dispatch** — `FW_ApplyTheme` iterates
  `oFrm.Controls` and dispatches on `ctrl.Tag`, not `ctrl.Name`.
  Adding a new control requires only a Tag in the designer.
- **`summary.hdr.models_full`** unified summary key. `with_idw` /
  `no_idw` reworded to "covered" / "missing".

### Changed

- **frmReport:** `lblSel` removed; all row details live in
  `lstReport` columns. `lblTitle` becomes static (`report.form.title`).
- **frmProgress:** no `FIT_FitLabel` on `lblFile`; designer font,
  AutoSize and WordWrap preserved.
- PRJ_VERSION bumped 9.18.8 -> 9.19.0.

### Test suite

- 25 tests, all passing.

## [9.18.8] - 2026-10-02

### Changed

- **frmReport:** `CKeyPreview` keyboard hook disabled on this form.
  The classes (`CKeyPreview`, `CKeyPreviewHook`) remain in the
  project for future use.
- PRJ_VERSION bumped 9.18.7 -> 9.18.8.

### Test suite

- 25 tests, all passing.

## [9.18.7] - 2026-10-02

### Fixed

- **Button captions now localized** (§18.1). Prior code never set
  the caption of any `lblXxxTxt` control — the designer default
  ("Run", "Cancel", "Browse...", etc.) stayed in place regardless
  of the `Lang` setting. The lang keys `btn.run`, `btn.cancel`,
  `btn.browse`, `btn.copy` existed in `modLang` but were never
  applied. New `PRJ_LocalizeCaptions` procedure (modUIHelpers)
  walks a per-form map and sets every button caption from its lang
  key. Called from `UserForm_Activate` of all three forms.
- **Filter button captions now localized** (§18.1). The filter
  buttons (`All`, `Errors`, `Skipped`, `OK`, `No .idw`) had no
  lang keys at all. Added `report.filter.*` keys to all 13
  languages and wired them into `PRJ_LocalizeCaptions`.
- **Form title now localized** (§18.1). Prior code used
  `m_sOrigCap` (designer caption = "Batch PDF Publisher") and
  fell back to `L("title")` only if the caption was empty — which
  it never was. Result: window always said "Batch PDF Publisher
  vX.Y.Z" even in German/Japanese/etc. Now uses `L("title")`
  directly.
- **Report header date format** (§18.2). `modReportBuild.RP_BuildHeader`
  used a hardcoded `"dd.MM.yyyy HH:mm"` format string instead of
  `L("report.date.format")`. Now uses the localized format.

### Added

- `modUIHelpers.PRJ_LocalizeCaptions` — set button and label
  captions from lang keys. Called from each form's
  `UserForm_Activate`.
- `modUIHelpers.SetCap` — private helper, sets caption from a lang
  key, silently skips missing controls.
- 10 new lang keys in all 13 languages:
  - `btn.back`, `btn.export`, `btn.retry`, `btn.reset`
  - `report.filter.all`, `report.filter.err`,
    `report.filter.skip`, `report.filter.ok`,
    `report.filter.noidw`
- ARCHITECTURE.md section 37 (Localization) documenting the
  externalization strategy, `PRJ_LocalizeCaptions`, language
  detection, and the A.66 exception for diagnostics.

### Changed

- PRJ_VERSION bumped from 9.18.6 to 9.18.7.
- ARCHITECTURE.md §21 (Hardcoded Strings Audit) updated to
  document `modSelfTest` diagnostic strings as an A.66 extension.

### Test suite

- 15 tests, all passing. No new tests — localization is a UI
  concern exercised manually by changing `Lang` in registry and
  restarting.
- Manual verification: set `Lang = "de"`, restart Inventor, open
  `frmScan`. All button labels ("STARTEN", "Durchsuchen...", etc.),
  filter labels, and the window title ("Stapel-PDF-Publisher
  v9.18.7") should be German.

## [9.18.6] - 2026-10-02

### Added

- `docs/ACCESSIBILITY.md` — comprehensive accessibility
  documentation. Covers WCAG 2.2 AA conformance detail (per success
  criterion), ISO/IEC 40500:2025, EN 301 549 V4.1.1 mapping,
  DOJ rule applicability, colour contrast table, keyboard
  navigation map, screen reader behaviour, and accepted limitations.
- `modSelfTest.PRJ_CheckAccessibility` — walks every form's controls
  and reports interactive controls missing `ControlTipText`,
  interactive controls with `TabStop = False`, and other
  accessibility gaps. Called by `PRJ_SelfTest`; available standalone.
- Focus ring on `txtPath` (frmScan) and `txtSearch` (frmReport).
  `Enter` / `Exit` handlers change `BorderColor` from `CLR_SEP` to
  `CLR_ACCENT` and back. Satisfies WCAG 2.2 SC 2.4.7 (Focus Visible).
- Tooltip on `txtSearch` in frmReport. Prior code set the tooltip
  only on `lblSearch` (the caption), leaving the actual input field
  without an accessible name source.
- ARCHITECTURE.md section 36 (Accessibility).

### Changed

- PRJ_VERSION bumped from 9.18.5 to 9.18.6.
- COMPLIANCE.md accessibility section replaced with a summary that
  points to `docs/ACCESSIBILITY.md`.
- README.md accessibility section updated.

### Fixed

- Screen reader users had no accessible name for `txtSearch`.

### Test suite

- 15 tests, all passing.

## [9.18.5] - 2026-10-02

### Added

- `LOG_CRITICAL` log level (value 4). `PRJ_Fatal` now emits at
  CRITICAL.
- Structured JSON logging (JSONL) via `LogJson` registry key.
- Multi-archive log rotation (5 archives retained).
- Dedicated audit file with 90-day retention.
- `PRJ_LogReloadLevel` and `PRJ_LogReloadJson` for runtime reconfig.
- ARCHITECTURE.md section 35 (Logging & Observability).
- CONTRIBUTING.md section "Logging".

### Changed

- `modLog.bas` restructured.
- Registry schema version 1 → 2.
- `STATE_ValidateConfig` validates `LogJson`.
- PRJ_VERSION bumped from 9.18.4 to 9.18.5.

### Fixed

- Single-archive log rotation could lose older diagnostic history.

### Test suite

- 15 tests, all passing.

## [9.18.4] - 2026-10-02

### Added

- `deps.lock.json` — human-editable dependency manifest.
- `modSelfTest.PRJ_CheckHostVersion`.
- ARCHITECTURE.md section 34 (Dependency Management).

### Changed

- `Build-SBOM.ps1` v2.2.0 → v2.3.0.
- PRJ_VERSION bumped from 9.18.3 to 9.18.4.

### Fixed

- Silent CISA 2026 violation in SBOM generation.

### Test suite

- 15 tests, all passing.

## [9.18.3] - 2026-10-02

### Added

- `STATE_MigrateSchema` in modState.
- `STATE_SCHEMA_VERSION` constant.
- ARCHITECTURE.md section 33 (Maintainability).
- ARCHITECTURE.md sections 16.4, 19.1, 29.14.

### Changed

- PRJ_VERSION bumped from 9.18.2 to 9.18.3.

### Test suite

- 15 tests, all passing.

## [9.18.2] - 2026-10-02

### Added

- Three new tests: `Test_PRJ_PathCombine`, `Test_RP_BuildSummaryText`,
  `Test_STATE_ValidateConfig`.
- Four new edge cases in `Test_SN_SanitizePath`.

### Changed

- Test count: 12 → 15.
- PRJ_VERSION bumped from 9.18.1 to 9.18.2.
- ARCHITECTURE.md gains section 32 (Testability).

### Fixed

- Historical inaccuracies in CHANGELOG (test count).

## [9.18.1] - 2026-10-02

### Fixed

- **frmScan reentrancy guards** — added `g_bInProcessLoop` checks to
  state-mutating handlers.
- **frmScan and frmProgress QueryClose** — X button converts to
  cancel request.
- **frmScan cancel button semantics** — sets `g_bCancel` when
  scanning.
- **Form-level keyboard handling** — replaced non-existent
  `KeyPreview` with `CKeyPreview` class.
- **frmReport reentrancy guards** — `g_bIsProcessing` checks.
- **modValidate.SC_Validate entry guard**.
- **modValidate duplicate SC_* helpers** removed.

### Added

- `CKeyPreview.cls`, `CKeyPreviewHook.cls`.
- ARCHITECTURE.md sections 31, 14.7.
- KNOWN_LIMITATIONS.md expanded concurrency section.

### Changed

- PRJ_VERSION bumped from 9.18.0 to 9.18.1.

### Test suite

- 12 tests, all passing.

## [9.18.0] - 2026-10-02

### Added

- Accessibility: SetFocus on txtSearch and lblCancel.
- Accessibility: ControlTipText for lblTitle, lblFill, lblBar, lblSearch.
- Reset confirmation dialog in frmScan.
- Checkbox accelerators S/F/A.
- ENTRY_* helpers, pre-flight TranslatorAddIn check.
- AUDIT log markers.
- CONTRIBUTING.md, COMPLIANCE.md, ROADMAP.md, TECH_DEBT.md.
- scripts/backup-ivb.ps1.
- Shell sanitization helper.
- modUIHelpers, modShellApp, PRJ_RunScanSession, STATE_ResetPreferences,
  STATE_AssertGlobals, RP_PickSummaryColor, RP_BuildSummaryText,
  RP_PickStatusBold, RP_GetSelectedPath, PRJ_ValidateFolderPath,
  SC_TickProgress, PRJ_OpenPathOrSelect, PRJ_PutClipboardFallback,
  PRJ_StartBatchFromScan.
- 8 named constants.
- 6 new lang keys in all 13 languages.
- 2 new tests: Test_STATE_AssertGlobals, Test_ERR_Classify.
- docs/ARCHITECTURE.md sections 7-24.

### Changed

- modEntry: Run_BatchPDFPublisher reduced to ~35 LOC.
- Thin wrappers removed.
- PRJ_CreatePDF split.
- Format -> Format$ in modFileInfo.
- Version sync to 9.18.0.

### Removed

- LOG_GetPath, SC_SetSum, SC_SetSum2.
- modNav, modReset, modTooltip, modBadge, modUI.
- RP_FormatSummary.

### Test suite

- Expanded from 10 to 12 tests.

## [9.17.0] - 2026-10-01

### Added

- Real-time progress in status line.
- Cancellation support during scan.
- Two-phase form activation.

### Fixed

- "Not Responding" state on large folders.
- ESC during scan.

### Technical

- SCAN_PROGRESS_MIN_INTERVAL_SEC (0.2s).
- g_bInProcessLoop scan-in-progress signal.
- 5 new lang keys.

## [9.16.6] - 2026-10-01

### Added

- Search box in report form.

## [9.16.5] - 2026-10-01

### Added

- Progress phase updates.

## [9.16.4] - 2026-10-01

### Changed

- Report export uses Shell.Application.ShellExecute.

## [9.16.3] - 2026-10-01

### Added

- Single-instance lock.
- g_bInProcessLoop guard claim (landed in v9.18.1).

## [9.16.2] - 2026-10-01

### Added

- Network failure detection.

## [9.16.1] - 2026-10-01

### Added

- Orphan cleanup handles .bak files.
- Antivirus hint.

## [9.16.0] - 2026-10-01

### Removed

- modAI (176 LOC dead code).

## [9.15.5] - 2026-10-01

### Changed

- TReportCounts Type.
- PDF_FinalizeExport dead param removed.

## [9.15.4] - 2026-10-01

### Changed

- FW_PickHoverColor.
- TPdfTiming Type.
- TBadgeCounts Type.

## [9.15.3] - 2026-10-01

### Added

- TStatus Enum + STS_Classify().

## [9.15.2] - 2026-10-01

### Added

- 8 new constants.

## [9.15.1] - 2026-10-01

### Changed

- Extracted palette to modPalette.

## [9.15.0] - 2026-10-01

### Added

- modTests - assertion-based test suite (5 tests).

## [9.14.8] - 2026-10-01

### Changed

- DST-safe timing via QueryPerformanceCounter.
- PRJ_CheckDiskSpace fail-open.
- SilentOperation sentinel.
- Log output ASCII-filtered.
- L() result cached in PROG_PickStatusColor.
- FILE_ExistsCached batch.
- Per-file MAX_PATH check.
- Pre-flight write access MsgBox.

## [9.14.7] - 2026-10-01

### Added

- UserForm_Terminate in all 3 forms.

### Changed

- PRJ_ShowScan error handler.
- QA_AnalyzeIDW per-op error handling.
- ReDim -> Collection in modSelfTest.
- InventorVersion check range.

## [9.14.6] - 2026-10-01

### Changed

- RP_FormatSummary and SC_BuildBadgeTooltip localized.
- PRJ_FormatDate uses L("report.date.format").
- 5 new lang keys.

## [9.14.5] - 2026-10-01

### Fixed

- frmScan non-existent controls crash.
- modReset RULE S.1 compliance.
- modDiagSimple RULE S.2 compliance.
- CACHE_Put* stamp guard.
- g_nTempCounter reset in STATE_Reset.
- g_RegCache CompareMode.
- L() fallback marker.
- LANG_SupportedCodes.

## [9.14.4] - pre-2026-10-01

Baseline. Version history before 9.14.5 is tracked in per-module
headers.

## Legacy versions

- v8.11 (VBA-DR-01) - earlier architecture.
- v1.7.0 (VBA-DR-01) - initial public-style release.