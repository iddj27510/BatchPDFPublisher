# ARCHITECTURE.md - VBA-DR-01 Batch PDF Publisher

**Version:** 9.25.42
**Host:** Autodesk Inventor 2027
**Runtime:** VBA 7.1
**Last updated:** 2026-10-04

Central architectural reference. Defines module responsibilities,
layer structure, dependency rules, invariants, state ownership,
validation layer, API isolation, circular dependency guards,
concurrency model, testability, maintainability, dependency
management, logging & observability, accessibility, localization,
and N/A items with future triggers.

---

## 1. Layers

The project is organized into 6 layers:

- **Core** - constants, state, entry point, navigation
- **Scan & PDF** - batch orchestration, PDF export pipeline, validation, inspection
- **Report** - report build, filter, export, summary aggregations
- **Infrastructure** - log, error, sandbox, COM, cache, file info, path, time, Win32, self-test
- **Localization** - 13-language strings and help content
- **UI** - forms, theme, hover, FIT, helpers, clipboard, browse, shell, keyboard hook

---

## 2. Module Responsibilities

### Core

| Module | Responsibility | LOC | Depends on |
|---|---|---|---|
| modConstants | All Public Const declarations | ~200 | (nema) |
| modState | Globals, registry, reset orchestration, schema migration | ~470 | modCOM, modSandbox, modLang, modLog, modConstants |
| modEntry | Entry point + navigation + batch start + host guard + log reload | ~150 | modState, modSandbox, modLog, modCOM, modPalette, modError, modBatch, modTime, modConstants |

### Scan & PDF

| Module | Responsibility | LOC | Depends on |
|---|---|---|---|
| modBatch | Batch orchestration + file collection helpers | ~200 | modPDFExport, modPDFProgress, modFileInfo, modSummary, modTime, modLog, modConstants |
| modPDFExport | PDF export pipeline (8 phases) | ~395 | modPDFFileIO, modPDFProgress, modQA, modInspect, modCache, modSandbox, modLang, modConstants |
| modPDFFileIO | Temp file lifecycle + atomic replace + AddIn lookup | ~215 | modRetry, modSandbox, modCOM, modLog, modConstants |
| modPDFProgress | Progress dialog updater + timing | 217 | modStatus, modLang, modTime, modState, modConstants |
| modInspect | IDW/PDF inspection (read-only) | 254 | modCOM, modCache, modFileInfo, modShellApp, modConstants |
| modValidate | Folder validation orchestration | ~380 | modUIHelpers, modTheme, modInspect, modBatch, modPath, modSummary, modTime, modLang, modConstants |
| modStatus | Canonical status classifier | 120 | modConstants |
| modQA | IDW/PDF QA validation | 111 | modConstants, modFileInfo, modLog |
| modRetry | File operation retry (Full Jitter) | 96 | modFileInfo, modLog, modTime, modWinAPI, modConstants |

### Report

| Module | Responsibility | LOC | Depends on |
|---|---|---|---|
| modReportBuild | Report text construction (pure) | 199 | modStatus, modLang, modFileInfo, modConstants |
| modReportExport | Report file export | 110 | modSandbox, modCOM, modShellApp, modLog |
| modReportFilter | Report classification, filtering, counting, summary | ~200 | modStatus, modLang, modFileInfo, modPDFExport, modCOM, modLog, modConstants |
| modSummary | Audit aggregations | 130 | modFileInfo, modPath, modCOM, modLang |

### Infrastructure

| Module | Responsibility | LOC | Depends on |
|---|---|---|---|
| modLog | Structured file logging + levels + rotation + audit + PII masking | ~500 | modSandbox, modCOM, modConstants, modState |
| modError | Error reporting + classification + crash report | ~330 | modLog, modLang, modSandbox, modCOM, modConstants |
| modSandbox | Sandbox lifecycle + path combine helper | ~180 | modCOM, modLog, modConstants |
| modCOM | Lazy COM singletons (FSO, Apprentice) | 123 | modLog, modConstants |
| modCache | Page/sheet count caches with FIFO eviction | 286 | modFileInfo, modLog, modConstants |
| modFileInfo | File metadata + value formatters + disk space | ~230 | modCOM, modShellApp, modLog, modLang, modConstants |
| modPath | Path sanitization pipeline | 195 | modConstants |
| modTime | Monotonic high-resolution clock | 97 | - |
| modWinAPI | Win32 API declarations (Sleep) | 33 | - |
| modSelfTest | Startup health checks + accessibility audit | ~450 | modConstants, modCOM, modSandbox, modFileInfo, modState |

### Localization

| Module | Responsibility | LOC | Depends on |
|---|---|---|---|
| modLang | 13-language flat key->string table | ~2720 | modConstants |
| modHelp | 13-language Apprentice Server help content | 1135 | modLang, modSandbox, modCOM |

### UI

| Module | Responsibility | LOC | Depends on |
|---|---|---|---|
| modTheme | Theme application + cycle + run-state + text alignment | 139 | modPalette, modHover, modConstants |
| modPalette | Color palette + dark mode + high contrast detect | ~180 | modConstants |
| modHover | Hover state machine + release | ~275 | modTheme, CHoverBtn, modConstants |
| modFIT | Fit-to-width label framework | 200 | modConstants |
| modUIHelpers | Tooltips + badges + sleep + align + accelerators + SC_ helpers + captions | ~510 | modLang, modLog, modTime, modWinAPI, modConstants |
| modClipboard | Clipboard write with 3-tier fallback chain | ~110 | modSandbox, modCOM, modShellApp, modUIHelpers, modLog, modConstants |
| modBrowse | Folder picker via Shell.Application | 98 | modShellApp, modLog, modCOM |
| modShellApp | Isolation layer for Shell "cmd.exe" / "explorer.exe" | ~110 | modCOM, modLog |
| CHoverBtn (class) | Hover handler for lblXxx + lblXxxTxt pair | 85 | modHover, modLog |
| CKeyPreview (class) | Form-level keyboard hook (VBA has no KeyPreview) | ~200 | CHookTextBox, CHookCheckBox, CHookListBox, CHookComboBox, CHookOptionButton |
| CHookTextBox (class) | WithEvents wrapper for MSForms.TextBox KeyDown | 51 | CKeyPreview |
| CHookCheckBox (class) | WithEvents wrapper for MSForms.CheckBox KeyDown | 42 | CKeyPreview |
| CHookListBox (class) | WithEvents wrapper for MSForms.ListBox KeyDown | 42 | CKeyPreview |
| CHookComboBox (class) | WithEvents wrapper for MSForms.ComboBox KeyDown | 42 | CKeyPreview |
| CHookOptionButton (class) | WithEvents wrapper for MSForms.OptionButton KeyDown | 42 | CKeyPreview |
| frmProgress (form) | Progress dialog | ~85 | modBatch, modTheme, modUIHelpers, modFIT, modLang, modTime, CKeyPreview |
| frmReport (form) | Report dialog | ~660 | modReportFilter, modReportBuild, modReportExport, modTheme, modHover, modFIT, modLang, modClipboard, modUIHelpers, modPDFExport, modFileInfo, modBatch, CKeyPreview |
| frmScan (form) | Scan dialog | ~370 | modValidate, modBatch, modState, modPath, modBrowse, modLang, modTheme, modHover, modFIT, modUIHelpers, modEntry, CKeyPreview |

### Test

| Module | Responsibility | LOC | Depends on |
|---|---|---|---|
| modTests | Assertion-based test suite (43 tests) | ~2350 | all |

---

## 3. Documented Exceptions

### modLang (~2720 LOC)

- **Reason:** Flat key->string table for 13 languages. No logic.
- **Alternative:** 13 modules per language -> 52 modules.
- **Decision:** Keep as one module. Grew to ~2720 in v9.25.42 with
  three new badge keys (to_create, to_recreate, force_suffix)
  across all 13 languages.

### modHelp (1135 LOC)

- **Reason:** Prose content for 13 languages. No logic.
- **Decision:** Keep as one module.

### modLog (~500 LOC)

- **Reason:** v9.18.5 added audit routing, JSON mode, multi-archive
  rotation, and runtime reload.
- **Decision:** Keep as one module.

### modSelfTest (~450 LOC)

- **Reason:** v9.18.6 added `PRJ_CheckAccessibility` (~130 LOC).
- **Decision:** Keep as one module.

### frmReport (~660 LOC)

- **Reason:** Complex form with list, filters, retry, export, search.
- **Decision:** Keep as one form.

### modTests (~2350 LOC)

- **Reason:** 43 test functions plus shared helpers. Test-only.
- **Decision:** Keep as one module. Excluded from production builds
  via `#Const PRODUCTION_BUILD`.

### modListBoxFit (~1600 LOC)
Reason: Universal ListBox geometry with GDI measurement, caching,
and calibration. Complex but isolated.
Decision: Keep as one module. Planned for removal in v10.0.0
(replaced by Label-grid).

MouseScroll (~1100 LOC)
Reason: MIT-licensed Cristian Buse implementation for mouse wheel
scroll on modal/modeless UserForms.
Decision: Keep until v10.0.0 (replaced by native subclassing).

modWinAPI (33 LOC)

- **Reason:** Shared Win32 declarations (Sleep) with 2 callers.
- **Decision:** Keep as one module.

### 39 deliverable modules (above BP1.4 range)

- **Reason:** 6 layers, 13 languages, PDF pipeline with 8 phases,
  batch orchestration, error infrastructure, UI helpers, keyboard
  hook, self-test, test suite.
- **Decision:** 39 modules is a deliberate architectural choice.

---

## 4. Reentrancy Guards

| Guard | Module | Cycle broken |
|---|---|---|
| m_bInEnsureSandbox | modSandbox | LOG_Open -> PRJ_EnsureSandbox -> PRJ_FSO recursion |
| m_bInCreateFSO | modCOM | PRJ_FSO -> PRJ_LogWarn -> LOG_Open -> PRJ_FSO recursion |
| PRJ_LogSubsystemBusy() | modLog (public) | modSandbox EH log while log mid-init |
| m_bInRotate | modLog | LOG_Rotate -> LOG_OpenInternal recursion |
| m_bInOpen | modLog | LOG_Open -> LOG_OpenInternal recursion |

---

## 5. Dependency Diagram

    frmScan
      -> modValidate -> modInspect -> modCOM
                     -> modUIHelpers
                     -> modSummary
                     -> modBatch
      -> CKeyPreview -> CHookTextBox / CHookCheckBox / CHookListBox
                     -> CHookComboBox / CHookOptionButton

    frmReport
      -> modReportFilter
      -> modReportBuild
      -> modReportExport
      -> modPDFExport
        -> modPDFFileIO -> modRetry
        -> modPDFProgress
        -> modQA
        -> modInspect
        -> modCache
      -> CKeyPreview -> CHook*

    frmProgress
      -> modBatch
      -> modPDFProgress
      -> CKeyPreview -> CHook*

    modEntry -> modState -> modConstants
            -> modBatch
            -> modSandbox
            -> modPalette

    modLog -> modSandbox -> modCOM -> modConstants
           -> modState

    modSelfTest -> modCOM
                -> modSandbox
                -> modFileInfo
                -> modState

---

## 6. Appendix A References

| Reference | Meaning |
|---|---|
| A.13 | Standardized header template |
| A.31 | Canonical status strings |
| A.61 | Multi-language help pattern |
| A.66 | English crash labels (debuggability) |
| A.96 | RULE S.2 (no decorative separator lines) |
| A.101 | Test suite + self-test |
| A.102 | Cache eviction pattern |
| A.103 | Hover state machine |

---

## 7. SRP Compliance

Every module has a single, clearly-defined responsibility. SRP test
applied to all 39 modules:

- Module has one reason to change.
- Module name describes its single responsibility.
- Module does not contain UI logic if logical module.
- Module does not perform API calls if UI module.
- Size is in 200-600 LOC range (or documented exception).

**Result:** 33 modules PASS without exception, 6 with documented
exception (modLang, modHelp, modLog, modSelfTest, frmReport,
modWinAPI, and 39-module count).

---

## 8. Thin Wrapper Audit

**Criterion:** Module is a "thin wrapper" if:
1. LOC < 100
2. No logic of its own (only delegates)
3. Another module is its natural host

**Consolidated in v9.18.0:**

| Module (pre) | LOC | Category | Action |
|---|---|---|---|
| modNav | 57 | Thin wrapper | Merged into modEntry |
| modReset | 70 | Thin wrapper | Merged into modState |
| modTooltip | 64 | Thin wrapper | Merged into modUIHelpers |
| modBadge | 113 | Thin wrapper | Merged into modUIHelpers |
| modUI | 146 | Thin wrapper | Merged into modUIHelpers |

**Kept as modules (not thin wrappers):**

| Module | LOC | Why it stays |
|---|---|---|
| modWinAPI | 33 | Shared Win32 declarations (documented exception) |
| modClipboard | ~110 | 3 fallback strategies with per-operation error handling |
| modBrowse | 98 | COM lifecycle + HWND ownership + CSIDL fallback |
| CHook* (5 classes) | 42-51 each | WithEvents requires a dedicated class per concrete control type |

**Result:** 0 thin wrappers in the project.

---

## 9. State Ownership

### 9.1 Two-tier state

**Tier 1 - Configuration layer (modState):**
All shared state lives in modState. This is the VBA equivalent of a
C# Singleton state class.

**Tier 2 - Private module state (m_*):**
Modules that need internal state declare Private m_* variables.

Rule: if state needs to be read/written by more than one module,
it belongs in modState.

### 9.2 Naming convention

| Prefix | Tier | Lives in | Example |
|---|---|---|---|
| g_ | Public global | Only modState | g_sLastFolder, g_colFiles |
| m_ | Private module | Any module | m_bOpen, m_dFreq |
| CLR_* | Public palette | modState | CLR_BG, CLR_ACCENT |
| UPPER_SNAKE | Const | modConstants | PRJ_VERSION, MAX_PATH_LEN |

### 9.3 Private state per module

| Module | Variables | Purpose |
|---|---|---|
| modLog | m_sLogPath, m_bOpen, m_nFileNum, m_bInRotate, m_bInOpen, m_bWriteFailed, m_aLogBuf(), m_nBufCount, m_sAuditPath, m_bAuditOpen, m_nAuditFileNum, m_bJsonMode, m_bJsonProbed | Log + audit + JSON state |
| modHover | m_sCurrentHover | Active hover button |
| modCOM | m_bInCreateFSO | Reentrancy guard |
| modSandbox | m_bInEnsureSandbox | Reentrancy guard |
| modPDFProgress | m_dLastDoEvents, m_oLblFile x 6 | Control cache |
| modTime | m_bFreqProbed, m_dFreq | QPC frequency cache |
| modLang | m_oDict, m_oEN, m_sLang | Language engine state |
| modState | m_oRegCache | Registry read cache |
| CKeyPreview | m_oForm, m_sCallback, m_colHooks | Keyboard hook registry |
| CHook* (5 classes) | m_ctl (WithEvents), m_oParent, m_sCallback | Single control subscription |
| frmScan | m_oKeyPreview, m_dToggleStartLeft, m_dToggleStartTop, m_dToggleHeight | Keyboard hook + toggle layout anchors |
| frmReport / frmProgress | m_oKeyPreview | Form-level keyboard hook instance |

### 9.4 Lifecycle

| Event | Action |
|---|---|
| Entry point | Run_BatchPDFPublisher: PRJ_IsHostAvailable, STATE_Init (runs STATE_MigrateSchema), CACHE_Clear, STATE_ValidateConfig |
| Batch start | PRJ_StartBatchFromScan does not reset state |
| Batch end | PRJ_RunScanSession sets g_dTimeEnd |
| Reset button | ResetClick calls STATE_ResetPreferences |
| Full reset | STATE_BackupPreferences + DeleteSetting + delete sandbox + STATE_Reset + STATE_Init |

---

## 10. Validation Layer

### 10.1 Four-level structure

| Level | Module | Responsibility | Type |
|---|---|---|---|
| 1. Input | modPath.SN_SanitizePath | Path format | Pure function |
| 2. Orchestration | modValidate.PRJ_ValidateFolderPath | Compose + resource check | Orchestrator |
| 3. Resource | modFileInfo.PRJ_CheckDiskSpace | Disk space | Win32 API |
| 4. Defensive | modBatch.PRJ_ScanFolder | Folder exists (redundant) | Guard |

### 10.2 Level 1 - Input validation

SN_SanitizePath checks: outer quotes, illegal chars, duplicate
backslashes, parent segments (capped at 20 iterations), UNC share
boundary escape, length (MAX_PATH_LEN = 260).

### 10.3 Level 2 - Orchestration

PRJ_ValidateFolderPath is the single entry point:
1. SC_QuickCheck (folder exists, length)
2. PRJ_CheckDiskSpace (if bRequireDiskSpace=True)

Called from frmScan.RunClick before batch start.

### 10.4 Level 3 - Resource validation

PRJ_CheckDiskSpace probes free disk space. Fail-open on probe error.

### 10.5 Level 4 - Defensive guard

PRJ_ScanFolder re-checks folder existence. Defense in depth.

### 10.6 Composition

    txtPath_Change -> SC_QuickCheck
    txtPath_AfterUpdate -> SC_Validate (full)
    RunClick -> PRJ_ValidateFolderPath(True)
      -> PRJ_StartBatchFromScan -> PRJ_ShowProgress
        -> PRJ_RunScanSession -> PRJ_ScanFolder (defensive)

---

## 11. Module Size Audit

### 11.1 Five categories

| Category | Range LOC | Action |
|---|---|---|
| Critical small | < 50 | Document or consolidate |
| Small | 50-199 | Document rationale |
| Target | 200-600 | No action |
| Tolerated | 600-800 | Document reason |
| Critical large | > 800 | Split or document exception |

### 11.2 Distribution

| Category | Count | % of total |
|---|---|---|
| Critical small (<50) | 2 | 5.1% |
| Small (50-199) | 18 | 46.2% |
| Target (200-600) | 15 | 38.5% |
| Tolerated (600-800) | 2 | 5.1% |
| Critical large (>800) | 2 | 5.1% |
| **Total** | **39** | 100% |

Napomena: `modTests` (~2350 LOC, test-only) nije uključen u gornju
distribuciju — on nije production modul.

### 11.3 Documented exceptions

- modWinAPI (33): shared Win32 declarations
- CHook* (42-51 each): WithEvents wrappers
- modLang (~2720): flat data table, no logic
- modHelp (1135): prose content, no logic
- modLog (~500): audit + JSON + rotation + reload
- modSelfTest (~450): +accessibility audit
- frmReport (~660): complex form
- modTests (~2350): test-only, excluded from production

### 11.4 Small module rationale

50-199 LOC range: legitimate SRP modules with 1-6 public functions.

### 11.5 Rule for future changes

New module <50 LOC: check if thin wrapper or documented exception.
50-199: document rationale. 200-600: no action.
600-800: document reason in section 3. >800: split or document.

---

## 12. UI / Logic Separation

### 12.1 Rule

Forms contain only UI logic:

**Allowed:** Event handlers, control property get/set, service calls,
local UI state.

**Not allowed:** Iterate business collections, compute business
values, direct Shell calls, direct registry access, direct API
calls, file I/O, business orchestration.

### 12.2 What was moved out

| Moved from | What | Where | Refactor |
|---|---|---|---|
| frmProgress | PRJ_ScanFolder call | PRJ_RunScanSession | 1.3 |
| frmReport | Retry loop (25 LOC) | RP_RetryPass | 1.3 |
| frmReport | Path parsing (25 LOC) | RP_GetSelectedPath | 1.3 |
| frmReport | Shell start/select | PRJ_OpenPathOrSelect | 1.3 |
| frmReport | Shell clip fallback | PRJ_PutClipboardFallback | 1.3 |
| frmReport | Summary counting | RP_CountFiles | 1.11 |
| frmReport | Summary text | RP_BuildSummaryText | 1.11 |
| frmReport | Summary color | RP_PickSummaryColor | 1.11 |
| frmReport | Status color inline | RP_StatusColor | 1.11 |
| frmReport | Status bold inline | RP_PickStatusBold | 1.11 |
| frmScan | REG_Set x 4 + state | PRJ_StartBatchFromScan | 1.3 |
| frmScan | Preference reset | STATE_ResetPreferences | 1.11 |
| frmScan | Folder validation | PRJ_ValidateFolderPath | 1.9 |
| All forms | Form-level keyboard hook | CKeyPreview class | 7.14 |
| All forms | Button captions | PRJ_LocalizeCaptions | 18.1 |
| frmScan | Badge composition | SC_BuildBadgeText (modUIHelpers) | 9.25.42 |
| frmScan | Toggle colour application | ApplyOneToggle (form-local helper) | 9.25.42 |

### 12.3 Current state

| Form | UI procedures | Business procedures |
|---|---|---|
| frmProgress | 10 | 0 |
| frmReport | 32 | 0 |
| frmScan | 28 | 0 |

---

## 13. API Isolation

### 13.1 Isolated APIs

| API | Accessor | Lives in |
|---|---|---|
| Scripting.FileSystemObject | PRJ_FSO() | modCOM |
| Registry | REG_Get / REG_Set | modState |
| Shell "cmd.exe" / "explorer.exe" | SHELL_* functions | modShellApp |
| Shell.Application.ShellExecute | SHELL_ShellExecute | modShellApp |
| Shell.Application factory | SHELL_GetShellApp | modShellApp |
| Sleep (Win32) | modWinAPI.Sleep | modWinAPI |
| QueryPerformanceCounter | TM_Now / TM_Elapsed | modTime |
| GetDiskFreeSpaceExW | PRJ_GetFreeSpaceBytes | modFileInfo |
| GetCurrentProcessId | Local in modLog, modPDFFileIO | Consumer |
| Path combine | PRJ_PathCombine | modSandbox |
| MSForms KeyDown (per control) | CKeyPreview.Initialize | CKeyPreview + CHook* |
| MSForms Enter/Exit (focus ring) | Control handler | frmScan / frmReport |

### 13.2 Partially isolated APIs

| API | Where | Why not isolated |
|---|---|---|
| Shell.Application.BrowseForFolder | modBrowse | Single caller |
| Shell.Application.Namespace | modInspect | Single use |
| WScript.Shell.RegRead | modPalette | Registry concern |
| ADODB.Stream | modReportExport, modSandbox | UTF-8 encoding |
| htmlfile | modClipboard | Clipboard, single use |
| OLE DataObject | modClipboard | Clipboard fallback |

### 13.3 Inventor Application - documented exception

ThisApplication is called from 8 modules. Each use is contextually
specific. Wrapping every call would create a 500+ LOC module with
40+ functions of different domains.

---

## 14. Circular Dependencies & Reentrancy Guards

### 14.1 Why cycles exist

Three subsystems (Log, Sandbox, COM) share resources and can be
called during each other's initialization.

### 14.2 Cycle 1 - Log to Sandbox to FSO (init)

    LOG_Open -> LOG_OpenInternal -> PRJ_EnsureSandbox -> PRJ_FSO
      -> (CreateObject fail) -> PRJ_LogWarn -> LOG_Write -> LOG_Open

**Guard:** m_bInOpen in modLog

### 14.3 Cycle 2 - Log to Sandbox (error path)

    LOG_Open -> PRJ_EnsureSandbox -> (CreateFolder fail) -> PRJ_LogError
      -> LOG_WriteError -> LOG_Write -> LOG_Rotate -> LOG_OpenInternal

**Guard:** m_bInEnsureSandbox + PRJ_LogSubsystemBusy()

### 14.4 Cycle 3 - FSO creation

    PRJ_FSO -> CreateObject -> (fail) -> PRJ_LogWarn -> LOG_Open
      -> PRJ_EnsureSandbox -> PRJ_FSO

**Guard:** m_bInCreateFSO in modCOM

### 14.5 Cycle 4 - Log rotation

    LOG_Write -> LOG_FlushBuffer -> LOG_Rotate -> LOG_FlushBuffer

**Guard:** m_bInRotate in modLog

### 14.6 Guards table

| Guard | Module | Cycle | Cleared |
|---|---|---|---|
| m_bInOpen | modLog | 1 | After LOG_OpenInternal |
| m_bInRotate | modLog | 4 | After LOG_Rotate |
| m_bInCreateFSO | modCOM | 3 | Before PRJ_LogWarn |
| m_bInEnsureSandbox | modSandbox | 2 | In EH and normal exit |
| PRJ_LogSubsystemBusy() | modLog | 2 | Public flag |

### 14.7 Form-level reentrancy guards

In addition to the initialization-order guards in 14.6, forms use
state flags to prevent handler reentry via the `DoEvents` message
pump. VBA is single-threaded, but `DoEvents` allows a message to
invoke another event handler while the current one is still on the
stack — the VBA equivalent of a race condition.

| Guard | Form / Module | Set by | Cleared by | Purpose |
|---|---|---|---|---|
| g_bInProcessLoop | frmScan, modValidate | SC_Validate entry | SC_Validate exit (normal and EH) | Prevents nested SC_Validate calls |
| g_bIsProcessing | frmReport, modEntry | RetryClick entry | RetryClick exit (normal and EH) | Prevents filter/list/copy/export during RP_RetryPass |
| m_bActivated | all forms | UserForm_Activate (first call) | UserForm_Terminate | Prevents duplicate initialization |
| m_bStarted | frmProgress | UserForm_Activate | — | Prevents re-entry on re-activate |
| m_bResetArmed | frmScan | ResetClick (first click) | ResetClick (second click), ResetDisarmCheck | Two-click confirmation |

**Rule:** any event handler that can be triggered from a message
pumped during a DoEvents block in another handler must check the
relevant guard at entry and exit silently if the guard is set.

See `modValidate.SC_Validate`, `frmScan.BrowseClick`, and
`frmReport.RP_SetFilter` for reference implementations.

**Historical note (v9.25.42):** the frmScan scan options used to be
CheckBoxes with an additional reentrancy guard `m_bInitializing`
(blocks Click during programmatic `.Value` assignment). That guard
was removed when the controls became Labels — Labels have no
`.Value` property, so no programmatic Click can be triggered.

---

## 15. Global Variable Audit

### 15.1 Philosophy

VBA has no Singleton pattern support. Module-level Public variables
in modState are the only way to share state. Project has 33 Public
g_* variables.

### 15.2 Categories

| # | Category | Count | Lifetime |
|---|---|---|---|
| 1 | Session lifetime | 10 | Whole session |
| 2 | Preferences | 3 | Persistent |
| 3 | Batch lifetime | 7 | One batch |
| 4 | Transient flags | 4 | One operation |
| 5 | UI state | 4 | Form/hover lifetime |
| 6 | Diagnostics | 2 | Error/abort |
| 7 | Counters | 3 | Various |

### 15.3 Inventory highlights

**Session (10):** g_sTheme, g_bThemeApplied, g_bCurrentDark,
g_bHighContrast, g_oFSO, g_oApprentice, g_dPageCache, g_dSheetCache,
g_nLogLevel, g_bAuditMode

**Preferences (3):** g_sLastFolder, g_bSubfolders, g_bForceOverwrite

**Batch (7):** g_colFiles, g_colLog, g_nTotal, g_nDone, g_dTimeStart,
g_dTimeEnd, g_sCurrentRelPath

**Transient (4):** g_bCancel, g_bInProcessLoop, g_bRetryMode, g_bIsProcessing

**UI (4):** g_colHoverBtns, g_bHoverRunEnabled, g_sActiveFilter, g_oProgress

**Diagnostics (2):** g_sCallStack, g_sAbortReason

**Counters (3):** g_nTempCounter, g_bAnalyzeIDW, g_bAnalyzePDF

### 15.4 Policy for new g_* variables

Only add when ALL THREE hold:
1. Cross-module
2. Cannot be a parameter
3. Documented in section 15.3

---

## 16. Public Variable Audit

### 16.1 Decision

Project does NOT replace Public variables with Property Get/Let.
This is deliberate, not technical debt.

Reasons:
1. VBA idiom
2. No safety gain
3. Performance
4. Readability
5. Refactor risk
6. Validation belongs at write-site

### 16.2 Inventory

- Public Const (~55): immutable, in modConstants
- Public CLR_* (~25): runtime palette, in modState
- Public g_* (33): session state, in modState
- Public Type (4): TReportCounts, TBadgeCounts, TPdfTiming, TStatus
- CHoverBtn.Public fields (7)

### 16.3 Policy

New Public variable:
- Immutable? -> Public Const in modConstants
- Runtime palette? -> Public CLR_* in modState
- Cross-module state? -> Public g_* in modState
- Single module? -> Private m_*
- Can be parameter? -> parameter

### 16.4 Invariants

STATE_AssertGlobals checks 8 invariants:
- g_sTheme in {I, S, L, D}
- g_nLogLevel in [LOG_DEBUG, LOG_CRITICAL]
- g_sActiveFilter non-empty
- g_nTotal >= 0
- g_nDone >= 0
- g_nDone <= g_nTotal (when g_nTotal > 0)
- g_dTimeEnd >= g_dTimeStart (when both > 0)
- g_nTempCounter >= 0

**Scope:** `STATE_AssertGlobals` is a test-only invariant checker.
It is called from `Test_STATE_AssertGlobals` in `modTests` and is
**not** invoked in production code. This is deliberate: production
invariant violations would fire on every batch and produce log
noise; the test suite is the right place to catch a broken
lifecycle before it reaches production. See
`ARCHITECTURE.md` section 32.1 (test suite scope).

---

## 17. Resource Separation

### 17.1 Five resource types

| Type | Location |
|---|---|
| UI strings | modLang |
| Colors | modState (CLR_*) + modPalette |
| Configuration | Registry (REG_Get/Set) |
| Help content | modHelp |
| Images | N/A (MSForms standard controls) |

### 17.2 UI strings - modLang

13 procedures: EN, DE, FR, IT, ES, PT, RU, PL, CS, JA, KO, ZH, ZHT.

API: L(key), L2(key, arg), L3(key, arg1, arg2), LN(key, args...)

Total: ~208 keys x 13 languages = ~2704 strings.

**No hardcoded UI strings** except in modError.PRJ_Fatal (A.66).

### 17.3 Policy

When adding any resource:
- UI string? -> modLang.L_FillXX in all 13 languages
- Color? -> modState.CLR_* + modPalette.FW_LoadPalette
- Config? -> modConstants.REG_KEY_* + REG_Get/Set
- Help? -> modHelp.HELP_* in all 13 languages
- Image? -> resources/ folder + LoadPicture

---

## 18. Naming Convention Deviations

### 18.1 Philosophy

OmniScope BP2.1: "Consistency is more important than perfect convention."

### 18.2 Deviation 1 - Control prefixes (§2.4)

Label-based buttons use lbl* prefix (not btn*). Semantic role in Tag.

### 18.3 Deviation 2 - Local camelCase (§2.7)

Hungarian-ish prefix style (sPath, bOK, nCount) - VBA idiom.

### 18.4 Deviation 3 - Boolean prefixes (§2.12)

b prefix (not is/has/can/should) - VBA idiom.

### 18.5 Policy

Forbidden: mixing styles, third prefix variant.
Required: new identifiers follow existing style.

---

## 19. Code Quality Audit

### 19.1 Section 3.1 Procedure Length

| Range | Procedures | Category |
|---|---|---|
| < 50 LOC | ~400 | PASS |
| 50-80 LOC | ~10 | Acceptable |
| 80-150 LOC | 0 | - |
| > 150 LOC | 0 | - |

**Cyclomatic Complexity (informal audit):**

| Procedure | CC (approx.) | Category |
|---|---|---|
| STS_Classify | 12 | Acceptable (<15) |
| SC_Validate | 10 | Acceptable |
| PDF_RunPipeline | 10 | Acceptable |
| RP_MatchesFilter | 9 | Target (<10) |
| SN_SanitizePath | 8 | Target |
| SC_ValidateDrawings | 8 | Target |
| SC_BuildBadgeText | 7 | Target |
| LOG_SanitizePII | 6 | Target |
| LOG_Asciify | 5 | Target |
| LayoutToggleLabels | 4 | Target |
| SN_ResolveParentSegments | 4 | Target |
| CACHE_EvictIfNeeded | 4 | Target |
| STATE_MigrateSchema | 4 | Target |
| PRJ_CheckAccessibility | 4 | Target |
| PRJ_CreatePDF | 3 | Target |

Threshold reference: target <10, warning >15, critical >25
(OmniScope §14.1). No procedure reaches the warning threshold.

VBA has no in-IDE CC measurement; these values are hand-counted
during code review. Rubberduck (VBE add-in) can automate this if
the project ever adopts it.

### 19.2 Section 3.2 Nesting Depth

Max depth = 3. No procedures above 3 levels.

### 19.3 Section 3.3 Duplicate Blocks

Extracted SC_TickProgress in modValidate.

### 19.4 VBA idioms (not duplicates)

Reentrancy guards, EH blocks, Set oX = Nothing, If oX Is Nothing,
Pattern B, OERN + Err.Number check.

---

## 20. Documentation Standards

### 20.1 Section 3.6 Mini-headers

Level 1 - Public API (XML-doc): '@Description, '@Param, '@Returns, '@Remarks
Level 2 - Private helper: one-line comment
Level 3 - Trivial: comment optional

### 20.2 Section 3.7 Header/Footer

Header format (required in every .bas and .cls file):

    Option Explicit
    ' Module    : <name>
    ' Purpose   : <one sentence>
    ' Reference : <Appendix or N/A>
    ' Version   : 9.25.42 (Inventor 2027)
    ' <changelog entries, newest first>
    '
    ' Requires  : <list of modules and functions>
    '
    ' Notes     : No line-continuation characters in this module.
    '             No decorative separator comment lines.

Footer format (required at end of every .bas and .cls file):

    ' END: <name> v9.25.42

UserForms (.frm) do not use header/footer in this format. They have
a VBE-generated header block plus Attribute declarations, followed
by the code. Footer is not applicable.

### 20.2.1 OmniScope Section 11.1 Field Mapping

OmniScope section 11.1 requires these header fields:
`NAZIV`, `TIP`, `OPIS`, `PROJEKAT`, `AUTOR`, `DATUM`, `VERZIJA`,
`SADRŽAJ`, `ZAVISNOSTI`, `UPOTREBA`.

The project uses an English-language short form that covers the
same information:

| OmniScope field | Project equivalent |
|---|---|
| NAZIV | `' Module    : <name>` |
| TIP | Inferred from file extension (.bas=Module, .cls=Class, .frm=UserForm) |
| OPIS | `' Purpose   : <one sentence>` |
| PROJEKAT | Project-wide constant `PRJ_VERSION` project name "VBA-DR-01" |
| AUTOR | Project-wide "portfolio-session" / Jugoslav Ilic (LICENSE, README) |
| DATUM | Project-wide release date 2026-10-04 (CHANGELOG, README) |
| VERZIJA | `' Version   : 9.25.42 (Inventor 2027)` |
| SADRŽAJ | `' #Region` markers in modules > 300 LOC (see 20.3) |
| ZAVISNOSTI | `' Requires  : <list>` |
| UPOTREBA | Documented in README.md + docs/ (see README quick start) |

Rationale for the deviation:
1. Per-module duplication of `PROJEKAT` / `AUTOR` / `DATUM` is
   noise — these are project-level constants, documented once in
   README.md and LICENSE.
2. `TIP` is redundant with the file extension (VBA convention:
   `.bas` = module, `.cls` = class, `.frm` = form).
3. `SADRŽAJ` is handled by `#Region` markers where useful
   (section 20.3), not as a static header block.
4. English field names align with the codebase language (all code,
   comments, and changelog entries are in English for portability).
5. VBA convention uses short `' Module :` / `' Purpose :` style —
   this matches existing VBE export headers.

All 10 OmniScope fields are traceable to a specific location. The
mapping satisfies §11.1 without adding 6 redundant lines to every
module header.

### 20.3 Section 3.8 Region Markers

Optional - used for modules > 300 LOC.

    ' #Region <section name>
    ...
    ' #End Region

### 20.4 Sections 11.4-11.6 - @Param, @Returns, @Throws Policy

OmniScope sections 11.4, 11.5, and 11.6 require that procedure
parameters, return values, and thrown exceptions be documented.

This project uses a **VBA-idiomatic subset** of those annotations:

| Annotation | Applied when |
|---|---|
| `@Description` | Public procedure with non-obvious purpose |
| `@Param` | Public procedure with 3+ parameters, or a parameter whose role is not evident from its name |
| `@Returns` | Public procedure whose return value has sentinel semantics (e.g. `-1` on miss, `""` on invalid input) |
| `@Throws` | **Never** — VBA does not have structured exception types. Every public procedure either returns a sentinel, sets a ByRef status, or delegates to `PRJ_ShowError`. There is no `throws` contract to document. |

Current state:
- `SN_SanitizePath` — full XML-doc (the reference example)
- `PRJ_CreatePDF`, `PDF_RunPipeline`, `STS_Classify`, `TM_Now`,
  `TM_Elapsed`, `STATE_MigrateSchema`, `PRJ_CheckHostVersion`,
  `PRJ_LocalizeCaptions` — one-line comment above signature
- All other public procedures — self-documenting names + `Requires`
  header block

Rationale: VBA tooling (VBE, Rubberduck) does not surface XML-doc
annotations to the caller. Writing `@Param` for every one of the
~400 procedures in the project produces a maintenance burden with
zero runtime or IDE benefit. The information is preserved in the
`Requires` block at the module level, which *is* visible in the
file header.

If the project migrates to .NET (see `ROADMAP.md`), full XML-doc
with `@param`, `@returns`, and `@exception` will be adopted because
the .NET toolchain actually surfaces those annotations.

---

## 21. Hardcoded Strings Audit

| Location | String | Reason |
|---|---|---|
| modError.PRJ_Fatal | "Fatal error in", "Crash report:", "See log:" | A.66 |
| modError.PRJ_ShowError.Safe | "Batch PDF Publisher" caption | L() may have failed |
| modHelp.APP_REG_PATH_2027 | "C:\Program Files\Autodesk\..." | Fixed |
| modHelp.APP_REG_PATH_2026 | "C:\Program Files\Autodesk\..." | Fixed |
| modHelp.APP_DOWNLOAD_URL | "https://www.autodesk.com/..." | Public URL |
| modSelfTest output | "[PASS]", "[FAIL]", "[WARN]", "[INFO]", descriptions | A.66 extension — diagnostics in English for debuggability. See §37.7. |
| modLog context strings | "LOG", "AUDIT", module names | A.66 extension — internal identifiers, not user-visible |

**A.66 extension:** Self-test and log diagnostic strings remain in
English regardless of the UI language. This is deliberate — logs and
diagnostics must be readable by support and developers who may not
speak the user's language. See §37.7.

---

## 22. Complex Expressions Audit

Project avoids complex logical expressions. Rule: if logical
expression requires mental parsing effort, split into multiple
lines with local variables.

---

## 23. TODO / FIXME Policy

TODO formats:
1. ' TODO: <description> (<ticket>)
2. ' TODO: <description>
3. ' NOTE: <decision> (deadline: <date>)

**Current state:** 0 TODO/FIXME comments.

---

## 24. Defensive Programming Audit

Rule 1: Check before use
Rule 2: Fail-fast with clear message
Rule 3: Validation at entry
Rule 4: Per-operation OERN
Rule 5: Default values

100% public functions have entry checks or documented assumptions.

---

## 25. Error Handling & Resilience

### 25.1 Try-Catch / On Error

Every procedure performing I/O, API call, or external resource
access has an error handler.

### 25.2 Log context

Format: <timestamp>  LEVEL  [Context]  Message [stack: ...]

**PII masking (v9.18.0):** LOG_SanitizePII masks %USERPROFILE% and
\Users\<name>\ patterns.

### 25.3 Global handler

ERR_GlobalHandler in modError.

### 25.4 User messages

Use PRJ_ShowError. For fatal: PRJ_Fatal writes crash report first.

### 25.5 Resource release

Set oX = Nothing in Cleanup label. Idempotent.

### 25.6 Recoverable vs non-recoverable

7 categories (EC_*):

| Category | VB Err.Number | Recoverable? |
|---|---|---|
| EC_NETWORK | 53, 68, 76 | Yes |
| EC_PERMISSION | 70, 75, -2147024891 | No |
| EC_DATA | 9, 13, 94 | No |
| EC_FILE_LOCK | 55, 58, 61 | Yes |
| EC_COM | 429, -2147467259 | No |
| EC_USER_CANCEL | 18 | Yes |
| EC_UNKNOWN | all others | No |

### 25.7 Inventor COM release

VBA does not support Marshal.ReleaseComObject. Use Set oX = Nothing.

### 25.8 Error strings

All user-visible error strings go through L(), L2(), L3(), LN().
Exceptions: A.66 crash report labels.

### 25.9 OERN scope

Minimal - around one operation, not a block.

### 25.10 Exception language

User-visible: localized. Internal log: EN.

### 25.11 Test coverage

43 tests, ~118 assertions.

### 25.12 Log level for fatal errors

`PRJ_Fatal` emits at `LOG_CRITICAL`, not `LOG_ERROR`. Fatal errors
are filterable separately from ordinary errors, both in plain-text
(`CRITICAL`) and JSON (`"level":"CRITICAL"`) modes. See §35.1.

---

## 26. Not Applicable Items (VBA Scope)

### 26.1 Philosophy

N/A items documented with reason + future trigger.

### 26.2-26.10 N/A Items

- §4.13 Circuit Breaker - N/A (no network services)
- §4.14 Background Threads - N/A (VBA single-threaded)
- §5.2 Connection Strings - N/A (no DB)
- §5.6 BinaryFormatter - N/A (.NET API)
- §5.7 CVE-2026-69439/71328 - N/A (.NET PDB)
- §5.8 CVE-2026-62886 - N/A (.NET Framework)
- §5.9 .NET Framework patch - N/A
- §5.13 Vault privileges - N/A (no Vault)
- §5.20 AI SBOM - N/A (no AI)

### 26.11 Summary

| Item | Category | Future trigger |
|---|---|---|
| Section 4.13 Circuit breaker | Network | Vault API integration |
| Section 4.14 Background threads | Runtime | .NET migration |
| Section 5.2 Connection strings | DB | DB integration |
| Section 5.4 Dependency verification | Supply chain | PASS via deps.lock.json + sbom.cdx.json (section 34) |
| Section 5.6 BinaryFormatter | .NET | .NET migration |
| Section 5.7 CVE PDB | .NET | .NET migration |
| Section 5.8 CVE .NET Framework EoP | .NET | .NET Framework (rejected) |
| Section 5.9 .NET Framework patch | .NET | .NET Framework (rejected) |
| Section 5.11 HTTPS | Network | No HTTP calls exist |
| Section 5.13 Vault privileges | Vault | Vault API integration |
| Section 5.14 XXE protection | XML | Future guidance in SECURITY.md |
| Section 5.18 Stale dependencies | Supply chain | Windows Update + Autodesk patches |
| Section 5.20 AI SBOM | AI | AI integration |
| Section 5.21 EU CRA | Compliance | SECURITY.md CSIRT SLA |

---

## 27. Performance & Memory Management

### 27.1 Section 6.1 Inventor API Batching - PASS

PDF_PrepareDocument calls Save2 only if oDoc.Dirty. Export uses
Publish_All_Sheets=True.

### 27.2 Section 6.2 Expensive Operations - PASS

Known expensive operations mitigated via caching, Apprentice-first
IDW sheet count, PDF page count cache.

### 27.3 Section 6.4 Loop Optimization - PASS

SC_TickProgress throttled to 0.2s. PDF_ThrottledDoEvents at 0.05s.

### 27.4 Section 6.5 String Concatenation - PASS

Array + Join in RP_BuildPdfDetail. LOG_FormatLine uses
concatenation but that runs once per log line, not in a hot loop.

### 27.5 Section 6.6 Lazy Loading - PASS

Caches with mtime stamps. Lazy singletons. Audit file opened lazily
on first AUDIT write.

### 27.6 Section 6.7 Progress Feedback - PASS

frmProgress + throttled updates.

### 27.7 Section 6.8 Background Threads - N/A

See section 26.2.

### 27.8 Section 6.9 COM Release - PASS

Set oX = Nothing pattern.

### 27.9 Section 6.10 Caching - PASS

| Cache | Max | Eviction |
|---|---|---|
| g_dPageCache | 10000 | FIFO batch 2000 |
| g_dSheetCache | 10000 | FIFO batch 2000 |
| m_oRegCache | unbounded | Never |
| FILE_FolderCache | per-folder | On change |
| m_bJsonProbed | once | Reload via PRJ_LogReloadJson |

### 27.10 Section 6.11 Memory Leak Fix

- PRJ_InitHover calls FW_ReleaseAllHoverButtons first
- UserForm_Terminate added to all 3 forms
- CKeyPreview.Terminate called from each form's UserForm_Terminate

### 27.11 Section 6.12 Excel - N/A

### 27.12 Section 6.13 Profiling - PASS

PDF_LogTiming per-phase timing.

### 27.13 Section 6.14 SQL - N/A

---

## 28. UI/UX Standards

### 28.1 Section 7.1 Visual Consistency - PASS

Segoe UI 9pt, 12px margins, 6px button gaps. Scan option toggles use
`TOGGLE_GAP = 6` between adjacent labels and `TOGGLE_WIDTH_PAD = 6`
inside each label.

### 28.2 Section 7.3 Control Dimensions - PASS

Button 72x24, filter 42x24, TextBox H=18/36, Label H=18/24.
Scan option toggles: `Height = designer height`, `Width = Ceiling
(natural width) + 6`.

### 28.3 Section 7.5 Palette - PASS

Palette in modPalette.FW_LoadPalette. OLE_COLOR (0x00BBGGRR).
Values documented with rendered #RRGGBB comments.

### 28.4 Section 7.6-7.8 Contrast - PASS

| Pair | Ratio |
|---|---|
| White on #0066CC | 5.6:1 |
| #F1F1F1 on #1E1E1E | 15.9:1 |
| #A6A6A6 on #1E1E1E | 8.1:1 |

### 28.5 Section 7.9 Light Theme - PASS

### 28.6 Section 7.10 Button Zones - PASS

### 28.7 Section 7.11 CTA - PASS

### 28.8 Section 7.12 Responsiveness - N/A

MSForms does not have AutoScaleMode.

### 28.9 Section 7.13 Tab Order - PASS

### 28.10 Section 7.14 Keyboard Navigation - PASS

| Form | Button | Accelerator |
|---|---|---|
| frmScan | Run | Alt+R |
| frmScan | Browse | Alt+B |
| frmScan | Theme | Alt+T |
| frmScan | Reset | Alt+E |
| frmScan | Cancel | Alt+C |
| frmScan | Subfolders toggle | Alt+S |
| frmScan | Force toggle | Alt+F |
| frmScan | Audit toggle | Alt+A |
| frmReport | Back | Alt+B |
| frmReport | Copy | Alt+C |
| frmReport | Retry | Alt+R |
| frmReport | Export | Alt+E |
| frmReport | Cancel | Alt+Q |
| frmReport | Filters | Alt+1..5 |
| frmProgress | Cancel | Alt+C |

ESC handled via CKeyPreview on all three forms (section 31.6).

### 28.11 Section 7.15 AccessibleName - N/A

MSForms limitation. `ControlTipText` is the fallback. See §36.

### 28.12 Section 7.16 Error Placement - Partial

lblSum near txtPath.

### 28.13 Section 7.17 Progress Feedback - PASS

### 28.14 Section 7.18 Destructive Actions - PASS

Two-click reset confirmation.

### 28.15 Section 7.19 High Contrast - PASS (v9.18.0)

PRJ_IsHighContrast via SystemParametersInfoW. FW_ApplyTheme returns
early if g_bHighContrast.

### 28.16 Section 7.20 WCAG 2.2 AA - PASS (within MSForms)

| SC | Status |
|---|---|
| 1.4.3 Contrast | PASS |
| 1.4.11 Non-text contrast | PASS |
| 2.1.1 Keyboard | PASS |
| 2.1.2 No keyboard trap | PASS |
| 2.4.3 Focus order | PASS |
| 2.4.7 Focus visible | Partial (TextBox only) |
| 3.2.1 On focus | PASS |
| 3.2.2 On input | PASS |
| 4.1.2 Name, role, value | N/A |

### 28.17 Section 7.21 EN 301 549 - Documented

See §36 and docs/ACCESSIBILITY.md.

### 28.18 Section 7.22 Decorative Separators - PASS

### 28.19 Section 7.23 Bold Usage - PASS

---

## 29. Data Management

### 29.1 Section 8.1 SQL - N/A

### 29.2 Section 8.2 DB Connections - N/A

### 29.3 Section 8.3 Config Model - PASS

Registry centralized in modState. 10 keys.

### 29.4 Section 8.4 Config Validation - PASS (v9.18.0)

STATE_ValidateConfig validates: Theme (ISLD), Sub (true/false),
Force (true/false), LogLevel (0-4), LogJson (true/false),
Lang (supported codes). Invalid values reset to defaults with
WARN log. v9.19.4 fixed the substring bug in STATE_ValidateEnum
(was accepting `"IS"` and `"ISL"` as valid theme codes).

### 29.5 Section 8.5 Temp Files - PASS

Cleanup on commit, on error, and at startup (>24h temp, >7 days bak).

### 29.6 Section 8.6 Config Backup - PASS (v9.18.0)

STATE_BackupPreferences exports all prefs to
%TEMP%\VBADR\prefs_backup_YYYYMMDD_HHMMSS.txt before reset.

### 29.7 Section 8.7 iProperties - N/A

### 29.8 Section 8.8 BOM - N/A

### 29.9 Section 8.9 Excel - N/A

### 29.10 Section 8.10 Validation - PASS

### 29.11 Section 8.11 File System Paths - PASS (v9.18.0)

PRJ_PathCombine in modSandbox is canonical path join. Uses
FSO.BuildPath when available.

Replaced in modLog, modError, modReportExport, modPDFFileIO,
modClipboard.

### 29.12 Section 8.12 XML - N/A

### 29.13 Section 8.13 BinaryFormatter - N/A

### 29.14 Registry schema versioning - PASS (v9.18.3)

See §33.5 for the schema migration contract.

---

## 30. Integration & API

### 30.1 Section 9.1 Inventor API Encapsulation - PASS

All Inventor API calls isolated in dedicated modules:

- modPDFExport, modPDFFileIO, modInspect, modCOM, modPalette,
  modEntry, modLang, modSelfTest

### 30.2 Section 9.2 Vault API - N/A

### 30.3 Section 9.3 Excel Interop - N/A

### 30.4 Section 9.4 Fallback Without Inventor - PASS (v9.18.0)

PRJ_IsHostAvailable in modEntry checks if ThisApplication is
accessible. Run_BatchPDFPublisher calls it first.

### 30.5 Section 9.5 Inventor 2027 API Novelties - Documented

See docs/INVENTOR_2027_NOTES.md. Summary:

| API | Status |
|---|---|
| SlotFeature | N/A |
| ReplaceBOMView | N/A |
| IrregularDistributionSet | N/A |
| GenerateOnSaveStatus | Future minor release |
| BrowserFolder | N/A |

### 30.6 Section 9.6 UseInventorAssemblyContext - Documented

For future compiled .ivb deployment. See INVENTOR_2027_NOTES.md
section 2.

### 30.7 Section 9.7 Autodesk Assistant - Documented

See INVENTOR_2027_NOTES.md section 3.

### 30.8 Section 9.8 Backward Compatibility - PASS

Tested on Inventor 2026 and 2027.

### 30.9 Section 9.9 API Timeouts - Documented

In-process COM. See INVENTOR_2027_NOTES.md section 5.

### 30.10 Section 9.10 API Versioning - PASS (v9.18.0)

PDF_FindAddIn in modPDFFileIO:

1. Primary: ApplicationAddIns.ItemById(PDF_GUID).
2. Fallback: enumerate, match by DisplayName containing "PDF".
   Logs WARN if fallback used.

PDF_ExportToTemp uses PDF_FindAddIn.

Host version guard: `HOST_IsSupportedMajor` in modConstants is the
single source of truth. `PRJ_CheckHostVersion` (modEntry) and
`PRJ_SelfTest` (modSelfTest) delegate to it. Accepts Inventor
Major 30 (2026) and 31 (2027).

### 30.11 Section 9.11 COM Event Deregistration - PASS

No COM event subscriptions. CHoverBtn and CHook* use MSForms
WithEvents (deregistered in Terminate).

### 30.12 Section 9.12 AddIn Manifest - N/A

VBA source distribution does not require manifest.

### 30.13 Section 9.13 Vault Fusion - N/A

### 30.14 Summary

| Item | Status | Where documented |
|---|---|---|
| 9.1 Inventor encapsulation | PASS | 13, 30.1 |
| 9.2 Vault API | N/A | 26.9 |
| 9.3 Excel Interop | N/A | 29.9 |
| 9.4 Fallback without Inventor | PASS | 30.4 |
| 9.5 Inventor 2027 novelties | Documented | INVENTOR_2027_NOTES |
| 9.6 UseInventorAssemblyContext | Documented | 30.6 |
| 9.7 Autodesk Assistant | Documented | 30.7 |
| 9.8 Backward compatibility | PASS | 30.8 |
| 9.9 API timeouts | Documented | 30.9 |
| 9.10 API versioning | PASS | 30.10 |
| 9.11 COM event dereg | PASS | 27.10, 30.11 |
| 9.12 AddIn manifest | N/A | 30.12 |
| 9.13 Vault Fusion | N/A | 30.13 |

---

## 31. Concurrency & Reentrancy

### 31.1 Platform constraints

VBA 7.1 in the Inventor host is single-threaded STA. There are no
threads, no async/await, no Task.Run, no locks, no thread-safe
collections. The only concurrency risk is **reentrancy**: a
`DoEvents` call pumps the Windows message queue, which may invoke
an event handler while the current handler is still on the stack.

Every "race condition" in this project is a reentrancy bug.

### 31.2 Reentrancy sources

`DoEvents` is called from four places:

| Location | Frequency | Impact |
|---|---|---|
| modValidate.SC_TickProgress | Throttled to 0.2s | During folder scan |
| modPDFProgress.PDF_ThrottledDoEvents | Throttled to 0.05s | During PDF export |
| modRetry.RETRY_Pause | Inside jitter sleep loop | During retry backoff |
| modUIHelpers.UI_SleepResponsive | Responsive sleep | During user feedback pauses |

Every one of these is inside a procedure that either modifies
shared state (`g_colFiles`, `g_sLastFolder`, `g_bSubfolders`) or
reads it in a tight loop.

### 31.3 Guards

VBA does not have a synchronization primitive. The project uses
state flags in `modState`:

| Flag | Owner | Lifetime | Consumers check at entry |
|---|---|---|---|
| g_bInProcessLoop | modValidate.SC_Validate | One validation run | frmScan event handlers, SC_Validate itself |
| g_bIsProcessing | frmReport.RetryClick | One retry pass | frmReport filter/list/copy/export handlers |
| g_bCancel | frmScan, frmProgress | One batch | Batch loop in modBatch |

### 31.4 QueryClose semantics

Forms that run modal work (`frmProgress`) or long-running
validation (`frmScan`) must handle the X button explicitly. Without
`UserForm_QueryClose`, clicking X during a batch unloads the form
mid-operation and leaves the batch running headless.

The pattern:

    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        If CloseMode = vbFormControlMenu Then
            Cancel = True
            g_bCancel = True
        End If
    End Sub

`CloseMode = vbFormControlMenu` distinguishes the user-initiated X
click from the programmatic `Unload Me` (which fires QueryClose with
`vbFormCode`). Programmatic unloads pass through unchanged.

### 31.5 Cancel button semantics

Cancel buttons in forms that host a batch or scan must NOT call
`Unload Me` directly. During a DoEvents pump, an event handler that
unloads its own form leaves the outer loop on the stack, running
against a terminated form instance.

Correct pattern (frmScan.lblCancel_Click):

    If g_bInProcessLoop Then
        g_bCancel = True
    Else
        Unload Me
    End If

The outer loop polls `g_bCancel` at the next iteration boundary and
exits cleanly. When idle, Cancel behaves as a normal close.

### 31.6 Form-level keyboard handling

VBA MSForms.UserForm does **not** have a `KeyPreview` property.
That is a Windows Forms (.NET) feature; FM20.DLL does not implement
it. There is no form-level `KeyDown` event that fires when a child
control has focus.

To get ESC handling regardless of which control has focus, the
project uses a `CKeyPreview` class module that subscribes to the
`KeyDown` event of every focusable control on the form and forwards
it to a form-level callback (`OnFormKeyDown`). The callback has this
signature:

    Public Sub OnFormKeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)

Setting `KeyCode = 0` inside the callback consumes the key, matching
the behaviour of `Handled = true` in .NET.

**v9.19.7 refactor:** the original design used a single
`CKeyPreviewHook` class with `Private WithEvents m_ctl As
MSForms.Control`. `MSForms.Control` is a base type that does not
expose `KeyDown`; VBA raised error 459 ("Object or class does not
support the set of events") the moment a concrete control was
assigned. The fix was one class per concrete control type:

| Class | WithEvents type |
|---|---|
| `CHookTextBox` | `MSForms.TextBox` |
| `CHookCheckBox` | `MSForms.CheckBox` |
| `CHookListBox` | `MSForms.ListBox` |
| `CHookComboBox` | `MSForms.ComboBox` |
| `CHookOptionButton` | `MSForms.OptionButton` |

`CKeyPreview.HookOne` dispatches on `TypeName(oCtl)` and
instantiates the matching class. Adding a new control type requires
one new class and one new `Case` — no other wiring changes.

All three forms instantiate `CKeyPreview` in `UserForm_Initialize`:

    Set m_oKeyPreview = New CKeyPreview
    m_oKeyPreview.Initialize Me, "OnFormKeyDown"

and release it in `UserForm_Terminate`:

    If Not m_oKeyPreview Is Nothing Then
        m_oKeyPreview.Terminate
        Set m_oKeyPreview = Nothing
    End If

See `CKeyPreview.cls`, the five `CHook*.cls` classes, and the
`OnFormKeyDown` procedures in each form.

### 31.7 Ctrl+Z in txtPath — accepted limitation

Ctrl+Z in an MSForms TextBox is handled internally by the Windows
Edit control. The VBA `Change` event does not fire for Ctrl+Z.
Reading `.Text` during `KeyUp` or `MouseDown` resets the Edit
control's undo buffer (breaks Ctrl+Z entirely).

Inventor VBA has no workaround: the `AddressOf` operator is not
supported (runtime error 35010), so Win32 `SetTimer` and Edit
subclassing are both unavailable. `Application.OnTime` exists only
in Excel and Word. MSForms 2.0 has no `Timer` control.

The project uses `RevalidatePathIfChanged` (frmScan), called from
`UserForm_MouseMove`. Any mouse movement over the form refreshes
`lblSum` and `lblSum2`. Pressing Enter in the textbox runs the full
validation immediately via `txtPath_KeyDown`.

**Residual limitation:** pure keyboard Ctrl+Z with no subsequent
mouse movement does not refresh the labels. Documented in
`docs/KNOWN_LIMITATIONS.md` §"Ctrl+Z in txtPath" and
`TECH_DEBT.md` won't-fix entry TD-W09.

### 31.8 What this is not

- **Not** a substitute for real async. Batch cancellation remains
  at file boundaries, never mid-operation.
- **Not** thread-safe. It is single-threaded by construction. The
  guards prevent logic errors, not data races (there are none).
- **Not** deadlock-free in a general sense — there are no locks to
  deadlock. The reentrancy cycles in section 14 are the closest
  equivalent.

---

## 32. Testability

### 32.1 Test suite scope

`modTests` contains 43 assertion-based tests covering:

| Area | Tests |
|---|---|
| Path sanitization | Test_SN_SanitizePath (12 assertions incl. edge cases) |
| Path sanitization (UNC) | Test_SN_SanitizePath_UNC |
| Path combination | Test_PRJ_PathCombine |
| File I/O | Test_PDF_SafeReplace_Logical, Test_PDF_ValidateHeader, Test_PDF_BuildTempPath |
| Cache | Test_CACHE_RoundTrip, Test_CACHE_CaseInsensitive, Test_CACHE_EvictIfNeeded |
| Status pipeline | Test_STS_Classify, Test_STS_Classify_Edge, Test_STS_Classify_ValidateStatus, Test_RP_StatusTag, Test_RP_MatchesFilter, Test_RP_StatusColor |
| Formatting | Test_PRJ_FormatElapsed, Test_PRJ_FormatFileSize, Test_PRJ_FormatDate, Test_PRJ_FormatElapsed_Edge, Test_PRJ_FormatDate_Edge, Test_FormatFileSize_Boundary, Test_FormatElapsed_Boundary |
| Localization | Test_L_FallbackChain, Test_RP_BuildSummaryText, Test_L_Placeholders, Test_U_UnicodeDecoder, Test_LANG_NormalizeCode, Test_LANG_SupportedCodes |
| Report filtering | Test_RP_CountFiles, Test_RP_PickSummaryColor, Test_RP_PickStatusBold, Test_RP_FindStat, Test_RP_CountErrorRows |
| Error classification | Test_ERR_Classify, Test_ERR_IsRecoverable_All, Test_ERR_DescribeClass_All |
| Configuration | Test_STATE_ValidateConfig |
| Global state | Test_STATE_AssertGlobals, Test_STATE_Init_ClearsRegCache |
| Retry | Test_RETRY_DeleteFile_MoveFile |
| QA | Test_QA_AnalyzePDF_TooSmall |
| Relative path | Test_PRJ_GetRelPath |
| Clipboard | Test_PRJ_PutClipboardFallback_EnsuresSandbox |

The suite is guarded by `#Const PRODUCTION_BUILD = False`. Release
builds set it to True to exclude `RunAllTests` from the compiled
project.

### 32.2 What is not covered (and why)

**Inventor API.** Reading `ThisApplication.Documents.ItemByName`,
calling `oDoc.Update`, driving `TranslatorAddIn.SaveCopyAs` — none
of this can be tested headless. VBA does not run without its host.
Integration testing requires a live Inventor instance and is done
manually via the T01-T11 fixtures.

**Forms.** UserForm event handlers, hover state, theme application,
label toggle layout — these require a visible form. UI-level
regressions (like the v9.18.1 reentrancy bugs, or the v9.19.8 /
v9.25.42 checkbox quirks) are not caught by the headless test suite.
Manual verification steps are documented in `KNOWN_LIMITATIONS.md`.

**Registry state mutations.** `Test_STATE_ValidateConfig` exercises
`STATE_ValidateConfig`, but the test restores original values. It
does not test the interaction between config changes and an active
session — that is a manual concern.

**CKeyPreview.** The class that hooks every control's `KeyDown`
event can only be tested with a real form. There is no test for the
forwarding logic; the pattern is verified by manual interaction
(press ESC with focus in each control type).

**STATE_MigrateSchema.** Lifecycle hook that runs once per session.
Testing it would require manipulating the real registry. Manual
verification: delete `HKCU\...\VBADR01\Prefs\SchemaVersion`, restart,
confirm the key is re-created with the current schema version.

**PRJ_LocalizeCaptions.** UI-visible output requires a real form.
Manual verification: set `Lang` to a non-English value, restart,
confirm button captions change language.

### 32.3 Mocking strategy (or lack thereof)

VBA has no dependency injection, no interfaces, and no mocking
framework. Test doubles are limited to:

- **Factory substitution.** `modCOM.PRJ_FSO()` and
  `modCOM.PRJ_GetApprentice()` are the closest thing to a seam.
  A test could `Set g_oFSO = Nothing` and force the lazy singleton
  to recreate — but cannot substitute a fake object without
  rewriting the factory.
- **File system.** Tests use `PRJ_SandboxPath()` for temporary
  files and clean up in `Cleanup:` labels. No filesystem mock.
- **Registry.** `Test_STATE_ValidateConfig` writes real registry
  values and restores them. This is intentional — it exercises the
  real path, not a mock.

Adding real mocking would require extracting interfaces for every
COM dependency and routing all calls through them. Estimated effort:
~400 LOC of new abstraction + rewrite of all consumers. Not
justified for a ~7500 LOC project.

### 32.4 Invariant coverage

The most load-bearing invariants are covered:

- `STS_Classify` precedence (Skipped Empty before Skipped) —
  Test_STS_Classify would fail if reordered.
- `SN_SanitizePath` UNC boundary escape — Test_SN_SanitizePath
  asserts empty return on escape.
- `RP_MatchesFilter` for all six filter modes — Test_RP_MatchesFilter.
- `STATE_AssertGlobals` 8 invariants — Test_STATE_AssertGlobals.
- `STATE_ValidateConfig` reset + restore — Test_STATE_ValidateConfig.
- `PRJ_PathCombine` empty base / empty child edge cases —
  Test_PRJ_PathCombine.
- `STATE_Init` clears the registry cache — Test_STATE_Init_ClearsRegCache.
- `PRJ_PutClipboardFallback` recreates the sandbox —
  Test_PRJ_PutClipboardFallback_EnsuresSandbox.

Non-covered invariants (documented here so future contributors know):

- `LOG_SanitizePII` masking patterns — no test.
- `PRJ_PathCombine` behaviour with UNC bases — only basic cases
  tested.
- `CACHE_EvictIfNeeded` FIFO behaviour at 10,001 entries — no test
  (would be slow and produces no value in CI).
- `CKeyPreview` dispatch chain — no test (requires a live form).
- `STATE_MigrateSchema` — no test (requires real registry writes).
- `PRJ_CheckHostVersion` — no test (requires a specific Inventor
  version).
- `PRJ_LocalizeCaptions` — no test (requires a live form).
- `SC_BuildBadgeText` / `SC_BuildBadgeTooltip` — no test (format
  depends on `g_bForceOverwrite` and requires a live form).

### 32.5 Running the suite

From VBA Immediate Window:

    RunAllTests

Output goes to the Immediate Window (`Ctrl+G`). No file output, no
JUnit XML, no CI integration — VBA is not amenable to it. Release
builds exclude the suite entirely.

### 32.6 Extending the suite

New test:

    Private Function Test_NewThing() As Long
        On Error GoTo EH
        Dim nFail As Long
        nFail = nFail + AssertEquals(Actual, Expected, "tag")
        If nFail = 0 Then Debug.Print "PASS [NewThing]"
        Test_NewThing = nFail
        Exit Function
    EH:
        Debug.Print "FAIL [NewThing] - " & Err.Description
        Test_NewThing = 1
    End Function

Register in `RunAllTests`. Update the test inventory comment in the
modTests header. Bump PRJ_VERSION (PATCH for coverage-only, MINOR
if a new public API is under test).

Rules:

- Every bug fix gets a regression test.
- Tests must be deterministic (no reliance on system time, network,
  or file system state outside the sandbox).
- Tests must clean up their own resources in `Cleanup:` labels.
- Tests that mutate registry or global state must save and restore.

---

## 33. Maintainability

### 33.1 Cyclomatic Complexity

Covered in section 19.1. Summary: no procedure exceeds the warning
threshold. The highest-complexity procedures are the two central
dispatch points (`STS_Classify` at ~12, `RP_MatchesFilter` at ~9)
and the two orchestrators (`SC_Validate` and `PDF_RunPipeline` at
~10). All four are documented in place.

### 33.2 Coupling (CBO)

Coupling Between Objects, measured as number of modules imported by
a module.

| Module | CBO | Category |
|---|---|---|
| frmReport | 12 | Above BP14.4 threshold (>5) |
| frmScan | 11 | Above BP14.4 threshold |
| modValidate | 10 | Above BP14.4 threshold |
| modEntry | 8 | Above BP14.4 threshold |
| modPDFExport | 7 | Above BP14.4 threshold |
| modReportFilter | 6 | Slightly above |
| modBatch | 6 | Slightly above |

**Why acceptable:** All seven are coordinator/orchestrator modules.
Their job is to call into many subsystems. Reducing their CBO would
require a mediator or controller layer — already tracked as TD-03
and TD-07 in `TECH_DEBT.md`, with an estimated effort of L and no
functional benefit.

The rest of the project is well under threshold: most utility
modules (modPath, modTime, modWinAPI, modStatus, modCache) depend on
0-2 other modules.

### 33.3 Change isolation

Every common change touches a bounded, predictable set of modules:

| Change | Modules touched |
|---|---|
| Add new language | modLang (+1 function), CONTRIBUTING.md header note |
| Add new status prefix | modConstants (+1 STAT_* const), modStatus (+1 branch), modReportFilter (+1 tag) |
| Add new filter mode | modReportFilter (+1 Case), frmReport (+1 button + handler), modLang (+1 tip key + 13 filter key entries), modUIHelpers (+1 SetCap map entry) |
| Add new theme | modPalette (+1 SetResting), modTheme (+1 cycle branch) |
| Add new registry key | modConstants (+1 REG_KEY_* const), modState (+1 validation branch) |
| Bump registry schema | modConstants (STATE_SCHEMA_VERSION), modState (STATE_MigrateSchema Case) |
| Add new test | modTests (+1 function + register in RunAllTests) |
| Add new dependency | deps.lock.json (+1 entry), sbom.cdx.json (regenerate) |
| Add new log level | modConstants, modLog |
| Add new scan option toggle | frmScan (+1 label in designer, +1 Case in ApplyToggleColors, +1 in LayoutToggleLabels, +1 in PRJ_LocalizeCaptions, +1 in PRJ_SetupAccelerators, +1 in PRJ_ApplyTooltips, +1 Click handler, +1 Opt*Click procedure), modUIHelpers (+3 map entries), modState (+1 g_b* global), modLang (13 language entries) |
| Add new keyboard-hooked control type | CKeyPreview (+1 Case in HookOne + 1 Case in HasKeyDown) + 1 new CHook* class |

No cross-cutting changes. Adding a language does not touch the
status pipeline. Adding a filter does not touch the PDF exporter.

### 33.4 Open/Closed Principle

The project is **closed-for-modification on data, not on code**.
Two central dispatch procedures are not OCP-friendly:

- `STS_Classify` — dispatch on 11 status prefixes. Adding a new
  prefix requires editing this procedure.
- `RP_MatchesFilter` — dispatch on 6 filter modes. Adding a new
  mode requires editing this procedure.

**Why this is acceptable in VBA:**
- VBA has no interfaces, no virtual dispatch, no strategy pattern
  in the C# sense. A dispatch table would move the condition from
  code to data, but the ordering of the conditions is load-bearing
  (Skipped Empty before Skipped, Error before Created). A table
  encoding order as row order is **harder** to maintain because the
  ordering constraint becomes silent.
- The dispatch sets are stable: 11 statuses since v9.19.6, 6 filters
  since v9.16.6. No history of churn.
- The alternative is more code, less clarity, and no measurable
  reduction in change risk.

**OCP elsewhere:** `modLang` is fully open — adding a language is
+1 private function and does not touch any other code.
`modPalette` accepts a new theme color as data.
`PRJ_LocalizeCaptions` uses a per-form map; adding a control to the
map does not require touching the procedure body.
`ApplyToggleColors` / `ApplyOneToggle` (frmScan) is the same
pattern for scan options.
`CKeyPreview.HookOne` (v9.19.7) is the same pattern for hookable
control types — a new type needs one new class + one Case.

### 33.5 Registry schema versioning

`REG_KEY_SCHEMAVERSION` tracks the shape of the registry tree
(currently `2`). `STATE_MigrateSchema` reads it on startup, called
from `STATE_Init`. The procedure handles four cases:

1. **Fresh install** (no version key) — sets version to current.
2. **Same version** — no-op.
3. **Older version** — applies migrations in order, bumps version,
   logs the transition to AUDIT.
4. **Newer version** (registry from a newer build) — logs WARN,
   leaves registry untouched, no downgrade attempted.

Adding a schema bump:

1. Bump `STATE_SCHEMA_VERSION` in `modConstants`.
2. Add a `Case <old-version>` block in `STATE_MigrateSchema`.
3. Write the migration logic (rename keys, set defaults, clear
   obsolete entries).
4. Add an entry to `CHANGELOG.md` under the release that ships it.

Schema history:

| Version | Added | Release |
|---|---|---|
| 1 | Base schema (Theme, Path, Sub, Force, Audit, LogLevel, SchemaVersion, Lang) | v9.14.x |
| 2 | `LogJson` key | v9.18.5 |

Before v9.18.3, `REG_KEY_SCHEMAVERSION` was written by
`STATE_BackupPreferences` but never read — the value was dead state.
A schema change would have silently invalidated user preferences.
This section and the accompanying procedure close that loop.

### 33.6 Onboarding time

Estimated onboarding time for a VBA-fluent developer:

- Day 1: read `README.md`, `ARCHITECTURE.md` §1-14, run `RunAllTests`,
  read `modTests.bas`
- Day 2: read `CONTRIBUTING.md`, pick a small `TECH_DEBT.md` item,
  make the change, run tests

Realistic onboarding: **1.5 to 2 days** for a VBA-fluent developer.
Not measured empirically — this is an estimate based on the
documentation surface.

### 33.7 SemVer application

`PRJ_VERSION` is a single source of truth in `modConstants.bas`.
Every release gets a `## [X.Y.Z]` block in `CHANGELOG.md`. Bump
rules documented in `CONTRIBUTING.md` §Versioning:

- **MAJOR** — breaking change (registry schema, status string
  contract, public API removal)
- **MINOR** — new feature, backward compatible
- **PATCH** — bug fix, refactor, documentation

Version string appears in 9 places (documented in CONTRIBUTING.md
§Release checklist). All must be updated together.

### 33.8 Refactoring safety

Refactoring confidence by area:

| Area | Confidence | Coverage |
|---|---|---|
| Pure logic (path, status, formatters) | High | 43 tests, ~118 assertions |
| Cache | High | Test_CACHE_RoundTrip, Test_CACHE_CaseInsensitive, Test_CACHE_EvictIfNeeded |
| Config | High | Test_STATE_ValidateConfig, Test_STATE_AssertGlobals |
| File I/O | Partial | Test_PDF_SafeReplace_Logical (temp file only) |
| Localization | Partial | Test_L_FallbackChain, Test_RP_BuildSummaryText |
| UI forms | None | Manual only |
| Inventor API | None | Manual only |
| Log subsystem | Partial | Manual |
| Accessibility | Advisory | PRJ_CheckAccessibility |

**Practical rule:** pure logic refactors freely; UI and COM
refactors require manual acceptance testing against T01-T11
fixtures. This is documented in `CONTRIBUTING.md` §Testing and
`ARCHITECTURE.md` section 32.2.

### 33.9 Technical debt visibility

`TECH_DEBT.md` is the register. Format:

- **Open items** — priority (P1/P2/P3), effort (XS/S/M/L/XL)
- **Resolved** — with version and date
- **Won't fix** — with reason

Each item is also referenced from the section that documents the
underlying design decision (e.g., TD-03 ↔ section 33.2, TD-16 ↔
this section). This cross-referencing ensures that a reader who
encounters a limitation finds the debt entry, and vice versa.

---

## 34. Dependency Management

### 34.1 Inventory

The project has four runtime dependencies. All are proprietary and
supplied by the OS or by Autodesk. There are no third-party open
source packages.

| Dependency | Version | Producer | Source | License |
|---|---|---|---|---|
| VBA | 7.1 | Microsoft | Windows built-in | Proprietary |
| Inventor API | 2026 / 2027 | Autodesk | Inventor install | Proprietary |
| stdole | 2.0 | Microsoft | Windows built-in | Proprietary |
| MSForms (FM20) | 2.0 | Microsoft | Microsoft Office | Proprietary |

Plus a set of Windows COM components used at runtime but not tracked
as explicit dependencies (Scripting.FileSystemObject,
Scripting.Dictionary, System.Collections.ArrayList, ADODB.Stream,
Shell.Application, htmlfile, OLE DataObject, WScript.Shell). Their
versions track Windows Update and are outside SBOM scope.

### 34.2 Lock file — `deps.lock.json`

`deps.lock.json` is the human-editable source of truth for the
dependency set. It lists each dependency with name, version, purl,
license, source description, and a `paths` array of candidate
install locations on the build machine.

The file is a lightweight equivalent of `packages.lock.json` in a
.NET project. VBA has no package manager, but the same principle
applies: dependencies should be pinned, reviewable in code review,
and diffable in Git.

Hashes are **not** stored in `deps.lock.json`. They are computed at
build time by `Build-SBOM.ps1` from the actual files on the build
machine. This avoids the false-confidence problem of hand-edited
hashes.

### 34.3 SBOM generation

`Build-SBOM.ps1` reads `deps.lock.json`, resolves each dependency's
path, computes SHA-256 hashes, and emits `sbom.cdx.json` in
CycloneDX 1.5 format.

Workflow:

1. Edit `deps.lock.json` to add or update a dependency.
2. Run `.\Build-SBOM.ps1` on a machine where all dependencies are
   installed.
3. Commit both `deps.lock.json` and the regenerated `sbom.cdx.json`.

If any dependency's path cannot be resolved (e.g., Inventor is not
installed on the build machine), the script **fails** unless
`-AllowMissingHashes` is explicitly passed. This is a deliberate
CISA 2026 compliance stance: a SBOM with `"content": "unavailable"`
is not compliant, so the build must not silently produce one.

### 34.4 CISA 2026 Minimum Elements

The generated SBOM covers all 23 required elements:

**SBOM Metadata (9 elements):**

| Element | Where |
|---|---|
| SBOM Author | `metadata.authors` |
| Timestamp | `metadata.timestamp` |
| Data Format Name | `bomFormat` = "CycloneDX" |
| Data Format Version | `specVersion` = "1.5" |
| Tool Name | `metadata.tools[0].name` |
| Tool Version | `metadata.tools[0].version` |
| SBOM Version | `version` = 1 |
| Generation Context | `metadata.component.properties[Generation_Context]` |
| SBOM Author Signature | `metadata.component.properties[SBOM_Author_Signature_SHA256]` |

**Component Data (8 elements):**

| Element | Where |
|---|---|
| Component Name | `components[i].name` |
| Component Version | `components[i].version` |
| Component Producer | `components[i].author` |
| Component Hash Value | `components[i].hashes[0].content` |
| Component Hash Algorithm | `components[i].hashes[0].alg` |
| Component License | `components[i].licenses[0].license.name` |
| Other Unique Identifier (PURL) | `components[i].purl` |
| Component Type | `components[i].type` |
| Component Coverage | `components[i].coverage` |

**Practices & Processes (6 elements):**

| Element | Where |
|---|---|
| Dependency Relationship | `dependencies[0].dependsOn` |
| Known Unknowns | `properties[cisa:known_unknowns]` |
| BOM Manifest Source | `metadata.component.properties[BOM_Manifest_Source]` |
| Distribution Model | `metadata.component.properties[Distribution_Model]` |
| Host Application | `metadata.component.properties[Host_Application]` |
| Runtime | `metadata.component.properties[Runtime]` |

CISA 2026 accepts SPDX or CycloneDX. This project uses CycloneDX
(JSON). SWID is explicitly disallowed by CISA 2026.

### 34.5 Transitive dependencies

The Inventor API is a facade over a large set of Autodesk COM
servers. Their specific versions are resolved at runtime by
whatever Inventor installation is present, and they cannot be
enumerated at build time without launching Inventor.

Rather than pretend to enumerate them, `deps.lock.json` declares
them under `knownUnknowns`:

- **transitive** — Inventor API transitive COM components
- **runtime-resolved** — PDF Translator AddIn GUID
- **platform-native** — Windows COM components

CISA 2026 accepts an explicit declaration of unknowns as an
alternative to enumeration. Declaring the gap is preferable to
filling it with speculation.

### 34.6 Update workflow

To update a dependency version:

1. Update `deps.lock.json` with the new version, purl, and source.
2. Run `.\Build-SBOM.ps1` on a build machine with the new version
   installed.
3. Confirm the SBOM diff shows the expected change.
4. Update `CHANGELOG.md` under the release that ships the change.
5. Run `PRJ_SelfTest` on the new host to confirm the version check
   passes.

To add a new dependency:

1. Add an entry to `deps.lock.json` with a non-empty `paths` array.
2. Run `.\Build-SBOM.ps1`.
3. If the path cannot be resolved on the build machine, either
   install the dependency or pass `-AllowMissingHashes` (only for
   development builds, never for releases).
4. Update `COMPLIANCE.md` dependency table.
5. Update `SECURITY.md` dependency vulnerability scanning section.
6. Update `INVENTOR_API_GOTCHAS.md` if the dependency is a COM
   component with known quirks.

### 34.7 Host version check

`PRJ_CheckHostVersion` (modEntry) delegates to
`HOST_IsSupportedMajor` in modConstants. The helper accepts
Inventor Major version 30 (2026) and 31 (2027). Inventor uses an
offset of 1996 for `SoftwareVersion.Major`: 2022 is 26, 2023 is
27, 2024 is 28, 2025 is 29, 2026 is 30, 2027 is 31.

Prior to v9.25.42, the project compared against 26 and 27 believing
those were Inventor 2026 and 2027. Every user on Inventor 2026 or
2027 saw the "not officially supported" warning at startup. The
`HOST_IsSupportedMajor` helper is now the single source of truth,
consumed by `PRJ_CheckHostVersion` and `PRJ_SelfTest`.

Tracking: TD-R25 (resolved in v9.25.42).

### 34.8 CVE response

The dependency set is Windows-native and Autodesk-hosted. There is
no package manager for automated CVE monitoring. `SECURITY.md`
documents the manual review process:

- Windows Update handles OS-level COM CVEs.
- Autodesk patches handle Inventor API CVEs.
- Monthly manual review of Microsoft Security Advisories.

`sbom.cdx.json` includes component hashes so that a downstream
consumer can detect which exact version of each dependency was
shipped in a given release.

### 34.9 Future-proofing

If the project ever migrates to .NET (see `ROADMAP.md` v10.x), the
deps.lock.json → SBOM workflow maps cleanly:

- `deps.lock.json` → `packages.lock.json`
- `Build-SBOM.ps1` → `dotnet sbom-tool generate`
- CycloneDX 1.5 → CycloneDX (same or newer spec version)

The pattern (human-editable lock file, machine-generated SBOM,
mandatory hashes) is identical.

---

## 35. Logging & Observability

### 35.1 Log levels

Five levels (OmniScope §16.2):

| Level | Value | Purpose |
|---|---|---|
| DEBUG | 0 | Development diagnostics |
| INFO | 1 | Normal flow |
| WARN | 2 | Recoverable anomalies |
| ERROR | 3 | Operation failed |
| CRITICAL | 4 | Session-threatening (only PRJ_Fatal) |

The `LogLevel` registry key sets the minimum level to emit. Levels
below the threshold are dropped before writing.

### 35.2 Log format

Default plain text:

    yyyy-mm-dd HH:mm:ss  LEVEL     [Context]  Message

`LogJson = True` switches to JSONL:

    {"ts":"2026-10-02 14:32:11","level":"INFO","ctx":"PRJ_CreatePDF","msg":"Created 12 pages"}

Both formats pass through `LOG_SanitizePII` and `LOG_Asciify`.
JSON mode additionally runs `LOG_JsonEscape`.

### 35.3 Rotation

Live log: `%TEMP%\VBADR\BatchPDF_Log_<PID>.txt`.

On exceeding `LOG_MAX_BYTES` (1 MB), rotates into numbered archives:

- `prev1_<PID>.txt` — most recent
- `prev2_<PID>.txt` — one rotation older
- ...
- `prev5_<PID>.txt` — oldest retained (LOG_ROTATE_MAX = 5)

Beyond slot 5, oldest is dropped. Total per-process storage bounded
by `LOG_MAX_BYTES * (LOG_ROTATE_MAX + 1)` ≈ 6 MB.

### 35.4 Audit trail

Every entry with context `"AUDIT"` writes to **two** places:

1. Main log (subject to LogLevel filter).
2. Separate audit file `%TEMP%\VBADR\BatchPDF_Audit_<PID>.txt`.

Audit records bypass the LogLevel filter.

Retention:

| File | TTL |
|---|---|
| Main log + archives | 24 hours |
| Audit file | 90 days |

Current AUDIT events:

- `STATE_ResetPreferences`
- `PDF_CommitReplace`
- `PDF_RollbackReplace`
- `PRJ_CleanupOrphans`
- `STATE_MigrateSchema` (schema bump only)

### 35.5 PII masking

`LOG_SanitizePII` masks before writing:

- `Environ("USERPROFILE")` → `<userprofile>`
- `\Users\<name>\` → `\Users\<user>\`

Applied to every line, in both plain text and JSON mode.

### 35.6 Non-ASCII handling

`LOG_Asciify` replaces every char > 126 with `?`, except:

- `\t`, `\n`, `\r` (pass through)
- Middle dot `·` (U+00B7) → `|`

Consequence: log files are pure ASCII, readable on any locale.
Localized strings in log messages are mangled — deliberate; logs
are diagnostics, not user-facing output.

### 35.7 Runtime reconfiguration

Two procedures reload config without restart:

- `PRJ_LogReloadLevel` — re-reads `LogLevel`
- `PRJ_LogReloadJson` — re-reads `LogJson`

Wrapped by `PRJ_ReloadLogLevel` in modEntry:

    PRJ_ReloadLogLevel

Prior to v9.18.5, both values were read once at STATE_Init.

### 35.8 What is not covered

**Remote log sink.** VBA has no async network I/O. Not implemented.

**Log querying.** Plain files. No search, no filter, no index.

**Structured metrics.** PDF_LogTiming is diagnostics, not a metrics
system.

**Structured audit for enterprise SIEM.** Audit file is per-process.
Enterprise SIEM ingest requires external shipping mechanism.

### 35.9 Configuration keys

| Key | Default | Validation |
|---|---|---|
| `LogLevel` | `1` (INFO) | Must be 0..4 |
| `LogJson` | `False` | Must be "true" / "false" |

Both live under `HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs`.

`STATE_ValidateConfig` resets invalid values to defaults with a WARN.

---

## 36. Accessibility

### 36.1 Scope

The full accessibility conformance documentation lives in
`docs/ACCESSIBILITY.md`. This section covers the architectural
integration points — where accessibility concerns are handled in
the codebase and what the platform allows.

### 36.2 Where accessibility is handled

| Concern | Handled by | Notes |
|---|---|---|
| Contrast | modPalette.FW_LoadPalette | Per-theme palette, contrast ratios documented in §28.4 |
| High contrast detection | modPalette.PRJ_IsHighContrast | Windows HC API + early return in FW_ApplyTheme |
| Keyboard accelerators | modUIHelpers.PRJ_SetupAccelerators | Called from every form's UserForm_Initialize |
| Form-level ESC | CKeyPreview class | VBA MSForms has no KeyPreview property; see §31.6 |
| Focus ring on TextBox | frmScan / frmReport Enter/Exit handlers | BorderColor change on focus |
| Tooltips (AccessibleName fallback) | modUIHelpers.PRJ_ApplyTooltips | Every interactive control gets one |
| Accessibility audit | modSelfTest.PRJ_CheckAccessibility | Advisory; runs in PRJ_SelfTest |
| Tab order | Designer TabIndex | Verified by PRJ_CheckAccessibility |
| Scan option toggle state | Colour (accent vs neutral) + Alt+ accelerators | Labels have no `.Value`; visual state is the accent background |

### 36.3 Platform limitations

MSForms (FM20.DLL) does not expose:

- `AccessibleName`, `AccessibleDescription`, `AccessibleRole` —
  no UI Automation properties
- Focus visuals on Label controls — Labels cannot receive focus
- `AutoScaleMode` — no DPI-independent layout
- Live region announcements — no dynamic content announcement
- A checked/unchecked state on Labels — the v9.25.42 scan option
  toggles are Labels and have no `.Value` property

These are accepted. Full impact analysis and workarounds are in
`docs/ACCESSIBILITY.md` §10.

### 36.4 Focus ring implementation

The two TextBox controls (`txtPath` in frmScan, `txtSearch` in
frmReport) render a visible focus ring via the `Enter` and `Exit`
events:

    Private Sub txtPath_Enter()
        Me.Controls("txtPath").BorderColor = CLR_ACCENT
    End Sub

    Private Sub txtPath_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Me.Controls("txtPath").BorderColor = CLR_SEP
    End Sub

This satisfies WCAG 2.2 SC 2.4.7 (Focus Visible) for those inputs.
The rest of the UI uses Label-based buttons, which cannot receive
focus, so they have no focus ring — Alt+ accelerators are the
primary keyboard affordance.

### 36.5 Tooltip coverage

Every interactive control receives a `ControlTipText` via
`PRJ_ApplyTooltips`. This is the fallback for `AccessibleName`
(screen reader users hear the tooltip when they hover the control).

`PRJ_CheckAccessibility` walks every form's controls and reports
any interactive control missing a tooltip. It is advisory — a
missing tooltip does not fail the self-test, but should be fixed
before release.

### 36.6 Scan option toggles

The v9.25.42 scan option toggles (`lblOptSub`, `lblOptForce`,
`lblOptAudit`) are Labels, not CheckBoxes. Accessibility impact:

- **Visual state:** accent background when ON, neutral when OFF.
  Visible to sighted users.
- **Screen reader:** announced as static text. No "checked"/"unchecked"
  state is announced (MSForms Labels have no `.Value`).
- **Keyboard:** Labels do not receive Tab focus. `Alt+S` / `Alt+F` /
  `Alt+A` accelerators are the primary keyboard affordance.
- **Trade-off accepted:** removing three MSForms checkbox quirks
  (first-click after programmatic assignment, programmatic-Click
  side effect, duplicate Click) was worth the accessibility
  reduction. Mitigation via tooltips and accelerators covers the
  practical impact.

### 36.7 What is not covered

- **Runtime text resize.** MSForms does not support user-adjustable
  text size. DPI scaling is provided by Windows at the OS level.
- **Screen reader role announcement.** A Label-based button is
  announced as static text, not as a button. No workaround in VBA.
- **Live status announcements.** `lblSum` changes are not announced.
  Users poll via Alt+ shortcuts.
- **Toggle state announcements.** No screen reader support for the
  scan option toggles' ON/OFF state.

### 36.8 Conformance statement

The add-in targets WCAG 2.2 Level AA. Per §3 of
`docs/ACCESSIBILITY.md`, all criteria that can be satisfied within
MSForms are satisfied. Remaining criteria are documented as accepted
limitations, with a migration path via §11 (Future Improvements) if
the project moves to .NET.

### 36.9 Accessibility audit API

`PRJ_CheckAccessibility` (in modSelfTest) walks every form's
controls and appends a report to `sReport` (ByRef String). It is
called from `PRJ_SelfTest` and is also available standalone:

    PRJ_CheckAccessibility

Checks performed:

1. Interactive controls (TextBox, CheckBox, OptionButton, ComboBox,
   ListBox, Label with `btn-*` tag) without `ControlTipText`.
2. Interactive controls with `TabStop = False`.

A missing tooltip on a decorative Label is not reported — only on
controls the user can actually interact with.

---

## 37. Localization

### 37.1 Scope

Full localization infrastructure: `modLang` (13 languages, flat
key→string table), `modHelp` (13 languages, multi-line prose), and
`PRJ_LocalizeCaptions` (runtime caption application).

### 37.2 Language detection

`LANG_Init` (modLang) determines the active language:

1. **Registry override** — `REG_KEY_LANG` if non-empty.
2. **Host auto-detect** — `ThisApplication.LanguageCode`.
3. **Fallback** — English.

`LANG_SetLanguage` forces a specific code. `LANG_SupportedCodes`
enumerates valid codes:

    "", "en", "de", "fr", "it", "es", "pt",
    "ru", "pl", "cs", "ja", "ko", "zh", "zht"

The empty string means "auto-detect".

### 37.3 String lookup

Four functions in modLang:

- `L(key)` — flat lookup, returns `??key??` if not found.
- `L2(key, arg)` — replaces `{0}`.
- `L3(key, arg1, arg2)` — replaces `{0}` and `{1}`.
- `LN(key, args...)` — replaces `{0}` .. `{N}`.

Fallback chain: if a key is missing in the active language, `L`
falls back to the English table. If missing in English too, returns
`??key??` — a visible marker that a key was not defined.

### 37.4 Caption application

Prior to v9.18.7, button and label captions were set by the designer
and never updated. The `btn.run`, `btn.cancel`, `btn.browse`,
`btn.copy` lang keys existed in `modLang` but were never applied.

v9.18.7 adds `PRJ_LocalizeCaptions` in modUIHelpers, called from
every form's `UserForm_Activate`. It walks a per-form map of
(control, lang key) pairs and sets the caption:

    Case "frmScan"
        SetCap oFrm, "lblRunTxt", "btn.run"
        SetCap oFrm, "lblCancelTxt", "btn.cancel"
        SetCap oFrm, "lblBrowseTxt", "btn.browse"
        SetCap oFrm, "lblResetTxt", "btn.reset"
        SetCap oFrm, "lblOptSub", "lbl.subfolders"
        SetCap oFrm, "lblOptForce", "chk.force"
        SetCap oFrm, "lblOptAudit", "chk.audit"

`SetCap` silently skips missing controls (defensive — a form that
does not have a control listed in the map will not error).

### 37.5 Form title

`UserForm_Initialize` sets the window title:

    sCap = L("title")
    Me.Caption = sCap & " v" & PRJ_VERSION

Prior to v9.18.7, the code used `m_sOrigCap` (designer caption =
"Batch PDF Publisher") and fell back to `L("title")` only if the
caption was empty. Since the designer sets a non-empty caption, the
fallback never fired and the window was always English.

### 37.6 Locale-aware formatting

`PRJ_FormatDate` (modFileInfo) uses `Format$(d, "Short Date")` and
`Format$(d, "Short Time")` — the Windows Region settings. Prior to
v9.19.2, it used `L("report.date.format")`, which was a per-language
hardcoded template that ignored the user's actual Windows locale.

`PRJ_FormatNow` provides the same formatting for the current
wall-clock time.

### 37.7 Diagnostics remain English (A.66 extension)

Log messages, self-test output, and error-dialog technical labels
remain English regardless of the UI language. Rationale:

- **Support** — support engineers read logs across a multi-language
  user base; logs must be searchable with English keywords.
- **Developer debugging** — developers may not speak every supported
  language; an untranslated log is more useful than one they cannot
  read.
- **Localization keys** — internal identifiers (`[PASS]`,
  `[CRITICAL]`, `EC_NETWORK`) are not translated.

This extends A.66 (English crash labels) to cover all diagnostics.
UI-visible strings (button captions, dialog text, error messages)
remain fully localized.

### 37.8 Where new strings go

When adding a user-visible string:

1. Add the key to `L_FillEN` in `modLang`.
2. Add the translation to every other `L_FillXX` (12 more).
3. Reference it via `L("key")`, `L2`, `L3`, or `LN` at the call site.
4. If the string is a caption, add it to `PRJ_LocalizeCaptions` map.

Never write a user-visible string literal in a form or module. The
only exceptions are the A.66 diagnostic strings documented in §21.

### 37.9 Adding a new language

1. Add a `L_FillXX` procedure to `modLang` (copy from `L_FillEN`).
2. Add the code to `LANG_Init` dispatch.
3. Add the code to `LANG_SupportedCodes`.
4. Add the code to the normalisation logic in `LANG_SetLanguage`.
5. Add the prose translations to `modHelp`.

No other module changes are needed — all UI strings go through `L()`.

### 37.10 RTL languages

Not supported. MSForms RTL handling is limited and no demand has been
identified. Adding an RTL language would require form redesign
(right-to-left layout, alignment mirror). Documented as out of scope
in `KNOWN_LIMITATIONS.md`.

### 37.11 Key inventory

Current key count: ~208 per language (13 languages = ~2704 entries).
Distribution:

| Category | Keys (approx.) |
|---|---|
| Buttons (btn.*) | 10 |
| Tab labels (tab.*) | 3 |
| Settings (settings.*) | 6 |
| Form labels (lbl.*) | 3 |
| Options (opt.*, chk.*) | 4 |
| Theme names (theme.*) | 3 |
| Messages (msg.*) | 12 |
| Log (log.*) | 3 |
| Scan summary (scan.*) | ~33 |
| Progress (progress.*) | 4 |
| Status (status.*) | 3 |
| PDF errors (pdf.*) | ~20 |
| Report (report.*) | ~50 |
| Reset (reset.*) | 10 |
| Error dialog (errdlg.*) | 5 |
| Tooltips (tip.*) | 20 |
| Report filters (report.filter.*) | 5 |
| Summary headers (summary.hdr.*) | 5 |

v9.25.42 added 3 new keys:
`scan.sum.to_create`, `scan.sum.to_recreate`,
`scan.sum.force_suffix`. Each translated into all 13 languages.

A full key dump can be produced by walking `m_oEN.Keys` in the
Immediate Window after `LANG_Init`:

    Dim k: For Each k In m_oEN.Keys: Debug.Print k: Next

(Not exposed as a public API — for diagnostic only.)

---

*End of ARCHITECTURE.md*