# Contributing to Batch PDF Publisher

Thank you for contributing. This document describes the conventions,
workflow, and quality bar for this project.

## Repository layout

    D:\www\VBA-DR-01\
    |-- src/                 VBA source (.bas, .cls, .frm, .frx)
    |-- docs/                Documentation and archives
    |   |-- legacy/          Historical artifacts (previous versions)
    |-- scripts/             PowerShell helper scripts
    |-- releases/            Versioned release artifacts
    |-- backups/             Timestamped .ivb snapshots (Git-ignored)
    |-- T01_HappyPath/       Test fixtures (real Inventor files)
    |-- T04_Recursive/
    |-- T05_Collision/
    |-- T07_ErrorHandling/
    |-- T08_EmptyFolder/
    |-- T10_MultiSheet/
    |-- T11_DWG/
    |-- README.md
    |-- CHANGELOG.md
    |-- CONTRIBUTING.md
    |-- LICENSE
    |-- deps.lock.json
    |-- Build-SBOM.ps1
    |-- sbom.cdx.json

### Dev-only modules (a_ prefix)

Two modules carry the `a_` prefix to sort at the top of the VBE
Project Explorer and signal that they are development-only:

- `a_modDiag` — OmniScope diagnostic dump tool. Project-agnostic,
  reusable across VBA projects. Never modified for a specific project.
  Excluded from Git and from src/.
- `a_modExport` — Project-specific helper. Provides
  `ExportAllComponentsToSrc` used to sync the VBE project to src/.
  Excluded from Git and from src/.

Both must be deleted from the .ivb before release builds.

## Language policy

All code, comments, commit messages, and documentation must be in English.

This includes:

- Variable names, function names, type names
- Inline comments and block comments
- Changelog entries in module headers
- Error messages (unless localized via L())
- Commit messages
- Documentation files (.md)
- SBOM descriptions and PURL identifiers

Exceptions:

- Localization string values (L_FillXX tables in modLang) are per-language
- User-facing UI strings go through L() lookup, not hardcoded
- Log messages must remain English regardless of UI language (Appendix A.66)

## Development workflow

### 1. Backup before editing

    cd D:\www\VBA-DR-01
    .\scripts\backup-ivb.ps1

This creates a timestamped copy of the active .ivb in backups\. Run it
before any non-trivial change.

### 2. Edit in the VBE

Open Inventor, press Alt+F11.

### 3. Export changed modules

Use the helper macro `Export_All` (in a_modExport) from
the Immediate Window:

    Export_All

This re-exports every component to src\. Do not export manually; the
helper is deterministic and overwrites cleanly.

### 4. Compile and test

In VBE:

    Debug -> Compile VBAProject

In Immediate Window:

    RunAllTests

Expected output:

    TOTAL: 43 / 43 passed, 0 failed

### 5. Update CHANGELOG.md

Add a new entry at the top:

    ## [9.25.32] - YYYY-MM-DD

    ### Added
    - ...

    ### Changed
    - ...

    ### Fixed
    - ...

Follow Keep a Changelog (https://keepachangelog.com/).

### 6. Update PRJ_VERSION

In modConstants.bas:

    Public Const PRJ_VERSION As String = "9.25.32"

Sync the same version string in:

- README.md
- docs/ARCHITECTURE.md
- docs/KNOWN_LIMITATIONS.md
- docs/ROLLBACK.md
- docs/INVENTOR_2027_NOTES.md
- docs/INVENTOR_API_GOTCHAS.md
- docs/ACCESSIBILITY.md
- COMPLIANCE.md
- SECURITY.md
- ROADMAP.md
- deps.lock.json
- Build-SBOM.ps1 (default parameter $Version)
- Any module header where you made functional changes

### 7. Regenerate the SBOM

    .\Build-SBOM.ps1 -Version 9.25.32

This reads deps.lock.json, resolves every dependency path, computes
SHA-256 hashes, and writes sbom.cdx.json. If any dependency cannot
be resolved, the script fails — see the Dependencies section below.

### 8. Commit

    git add src/ docs/ CHANGELOG.md README.md deps.lock.json Build-SBOM.ps1 sbom.cdx.json
    git commit -m "feat: add <feature> (v9.25.32)"

Conventional Commits prefixes:

- feat:     new feature
- fix:      bug fix
- docs:     documentation only
- refactor: code change without behavior change
- test:     test additions or fixes
- chore:    build, tooling, dependencies
- perf:     performance improvement

### 9. Tag a release

    git tag -a v9.25.32 -m "Release v9.25.32"
    git push --tags

## Naming conventions

### Files and components

| Type      | Prefix       | Example        |
|-----------|--------------|----------------|
| Module    | mod*         | modState       |
| Class     | C* or cls*   | CHoverBtn      |
| Form      | frm*         | frmScan        |
| Constant  | UPPER_SNAKE  | MAX_PATH_LEN   |

### Controls

| Control    | Prefix |
|------------|--------|
| Label      | lbl*   |
| TextBox    | txt*   |
| CheckBox   | chk*   |
| ComboBox   | cbo*   |
| ListBox    | lst*   |
| Frame      | fra*   |
| ProgressBar| prg*   |

### Variables

Hungarian prefix is the project-wide convention:

- `s` = String
- `b` = Boolean
- `n`, `l` = Long
- `d` = Double, Date
- `o` = Object
- `v` = Variant
- `a` = Array
- `col` = Collection or ArrayList
- `t` = Type (UDT)
- `m_` = module-private field
- `g_` = global (only in modState and modPalette)

Examples:

    Private m_sOrigCap As String
    Private m_bStarted As Boolean
    Public g_sTheme As String
    Dim sPath As String
    Dim bForce As Boolean
    Dim nCount As Long
    Dim oDoc As Object

### Procedures

Module-specific prefixes are used consistently:

- `PRJ_*`  — public cross-module API
- `SC_*`   — frmScan scope and scan-time UI helpers
- `RP_*`   — frmReport scope
- `PROG_*` — frmProgress scope (progress dialog)
- `FW_*`   — framework (theme, hover)
- `CACHE_*`, `LOG_*`, `REG_*`, `STATE_*`, `TM_*`, `LANG_*`, `L_*`,
  `UI_*`, `SN_*`, `RETRY_*`, `PDF_*`, `QA_*`, `HELP_*`, `ENTRY_*`,
  `STS_*`, `PAL_*`, `TH_*`, `FIT_*`, `SHELL_*`, `DDS_*` (a_modDiag only)
  — module-specific prefixes

Procedure names use `prefix_Snake_Case` for the prefix part and
`PascalCase` for the rest: `PRJ_CreatePDF`, `SC_RunClick`,
`FW_ApplyTheme`, `CACHE_EvictIfNeeded`.

Event handlers follow `Source_Event`: `UserForm_Initialize`,
`lblCancel_Click`, `txtSearch_Change`.

Form-level keyboard callback follows `On<Scope><Event>`:
`OnFormKeyDown`.

## Code style

### RULE S.1 — No line continuations

Never use the underscore at end of line. Break long expressions into
intermediate variables.

BAD:

    sResult = L("key.a") & L("key.b") & _
              L("key.c") & L("key.d")

GOOD:

    sResult = L("key.a") & L("key.b")
    sResult = sResult & L("key.c") & L("key.d")

### RULE S.2 — No decorative separators

No `' =====`, `' -----`, `' *****` comment banners.

Section headers inside a module are allowed with a single-line
comment plus a blank line:

    ' Label-pair click handlers (Pattern B)

    Private Sub lblCancel_Click()

### Error handling

Every procedure that touches I/O, COM, or external resources must
have an `On Error GoTo` handler:

    Public Sub PRJ_DoWork()
        On Error GoTo EH
        ' ...
        Exit Sub
    EH:
        PRJ_LogError "PRJ_DoWork", Err.Description
    End Sub

`On Error Resume Next` must be scoped per-operation:

    On Error Resume Next
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        Err.Clear
        ' handle failure
    End If
    On Error GoTo 0

Never blanket-OERN across an entire procedure.

### Log before UI

Errors are logged before showing dialogs:

    PRJ_LogError "context", "message"
    PRJ_ShowError "context", "message"

This guarantees that a log entry exists even if the dialog is
dismissed or the session crashes.

### Comments

Comments explain WHY, not WHAT. Code shows what.

GOOD:

    ' CompareMode must be set before the first Add.
    ' Windows paths are case-insensitive.

BAD:

    ' Increment i by 1.
    i = i + 1

### XML doc for public procedures

Public procedures should carry XML doc annotations:

    '@Description "Sanitize a user-provided path for FSO use."
    '@Param "sIn" "Raw path string."
    '@Returns "Sanitized path, or empty string if invalid."
    Public Function SN_SanitizePath(sIn As String) As String

Private procedures may use a one-line comment above the Sub/Function.

## Pattern B — label pair event handling

Every clickable button in the UI is a Label pair:

- `lblRun`      — frame/background that receives mouse events
- `lblRunTxt`   — text label (child of the frame)

Both labels delegate to a named action sub:

    Private Sub lblRun_Click()
        SC_RunClick
    End Sub

    Private Sub lblRunTxt_Click()
        SC_RunClick
    End Sub

Never put logic in the `_Click` handler itself. Always delegate.

Rationale: one named sub per action means behavior is identical
regardless of which label was clicked, and the action can be invoked
from ESC handlers, keyboard shortcuts, and code without duplicating
logic.

## Pattern C — form-level keyboard hook

VBA MSForms does not have a `KeyPreview` property (that is a
Windows Forms .NET feature). To handle ESC at form level, forms use
the `CKeyPreview` class module that subscribes to `KeyDown` on every
focusable control and calls a form-level callback.

In `UserForm_Initialize`:

    Set m_oKeyPreview = New CKeyPreview
    m_oKeyPreview.Initialize Me, "OnFormKeyDown"

The callback:

    Public Sub OnFormKeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
        If KeyCode = 27 Then
            ' handle ESC
            KeyCode = 0   ' consume
        End If
    End Sub

In `UserForm_Terminate`:

    If Not m_oKeyPreview Is Nothing Then
        m_oKeyPreview.Terminate
        Set m_oKeyPreview = Nothing
    End If

Never attempt `Me.KeyPreview = True` — it does not exist in MSForms
and causes a compile error.

## Pattern D — single-label toggle

Since v9.25.32, scan options (`Include subfolders`, `Force recreate`,
`Audit drawings`) are single Labels, not CheckBoxes. Each label
carries both the background colour and the caption:

    lblOptSub     ' background + caption, active/inactive colour
    lblOptForce
    lblOptAudit

State is carried by a `g_b*` global flag. Colour is applied by
`ApplyToggleColors` (called after every Click and after every theme /
language refresh):

    Private Sub ApplyOneToggle(ByVal sName As String, ByVal bActive As Boolean)
        ' Active   -> CLR_ACCENT background, CLR_ACCENT_TXT text
        ' Inactive -> CLR_BTN_SEC background, CLR_BTN_SEC_TXT text
    End Sub

Layout is managed by `LayoutToggleLabels`. Order-critical quirk:
`AutoSize = False` must be assigned **before** `.Width` / `.Height`
because MSForms ignores those assignments while `AutoSize` is True.

Accelerators: `Alt+S`, `Alt+F`, `Alt+A` (Labels do not receive Tab
focus in MSForms).

Why single Label instead of CheckBox: three MSForms quirks disappear
at once (first-click after programmatic assignment, programmatic-
Click side effect, duplicate Click). See `CHANGELOG.md` v9.25.32.

## Testing

Tests live in modTests. Add a test for every bug fix.

Test template:

    Private Function Test_NewFeature() As Long
        On Error GoTo EH
        Dim nFail As Long
        nFail = 0

        nFail = nFail + AssertEquals(Actual, Expected, "tag")

        If nFail = 0 Then Debug.Print "PASS [NewFeature]"
        Test_NewFeature = nFail
        Exit Function
    EH:
        Debug.Print "FAIL [NewFeature] - " & Err.Description
        Test_NewFeature = 1
    End Function

Register the new test in `RunAllTests`. Rules:

- Every bug fix gets a regression test.
- Tests must be deterministic.
- Tests must clean up their own resources in `Cleanup:` labels.
- Tests that mutate registry or global state must save and restore.

See `ARCHITECTURE.md` section 32 for the testability strategy.

## Logging

Logging goes through `modLog`. Direct file writes are not permitted
in any other module. Every message passes through `PRJ_Log*` wrappers.

### API

    PRJ_LogDebug    "Context", "Message"
    PRJ_LogInfo     "Context", "Message"
    PRJ_LogWarn     "Context", "Message"
    PRJ_LogError    "Context", "Message"
    PRJ_LogCritical "Context", "Message"

The `Context` string identifies the caller. Convention:

- Module-level: `"modName"`
- Procedure-level: `"modName.Procedure"` or `"Procedure"` for
  cross-module entry points
- Special: `"AUDIT"` for compliance-relevant events (see below)

### Choosing a level

- **DEBUG** — development diagnostics. Not shown at default
  `LogLevel=1`. Use for tracing internal state during development.
- **INFO** — normal flow. Session start/end, batch start/end, file
  processed successfully, config reloaded.
- **WARN** — recoverable anomalies. Cache miss, fallback to slower
  path, invalid registry value reset to default. The operation
  completes but something unexpected happened.
- **ERROR** — operation failed. File export failed, network share
  unreachable, PDF header validation failed. The batch continues.
- **CRITICAL** — session-threatening. Only used by `PRJ_Fatal`.
  Startup failure, unrecoverable crash.

### PII

Never log raw `%USERPROFILE%` paths, usernames, or any value that
could identify the operator. `LOG_SanitizePII` masks `%USERPROFILE%`
and `\Users\<name>\` automatically, but it only catches those two
patterns. Do not log registry keys that hold user-specific
configuration, network share credentials, or anything else that
would not survive a security review.

### Audit events

Events that must be traceable for compliance go through
`PRJ_LogInfo "AUDIT", "..."`. The string `"AUDIT"` as the context
value triggers:

1. Normal log entry (subject to `LogLevel` filter).
2. Mirror write to the separate audit file with 90-day retention.

Current AUDIT events:

- `STATE_ResetPreferences` — user reset all preferences
- `PDF_CommitReplace` — backup deleted after successful PDF replace
- `PDF_RollbackReplace` — rollback on page mismatch
- `PRJ_CleanupOrphans` — orphan file cleanup at startup

New AUDIT events require justification in the PR description:
"which compliance or operational process needs this on the record?"

### Structured logging

The `LogJson` registry key switches between plain text and JSONL
output. When adding a log line, do not assume either format — the
`LOG_FormatLine` helper handles both. Just pass the level, context,
and message.

### Testing logging

Logs go to `%TEMP%\VBADR\BatchPDF_Log_<PID>.txt`. Audit records go
to `%TEMP%\VBADR\BatchPDF_Audit_<PID>.txt`. Neither is
production-monitored; open them in a text editor to verify.

To change log level without restarting Inventor:

    PRJ_ReloadLogLevel

See `ARCHITECTURE.md` section 35 for the full logging contract.

## Dependencies

The project has no package manager. Runtime dependencies are Windows
built-ins or Autodesk-supplied COM components, declared in
`deps.lock.json`. Hashes are computed at build time by
`Build-SBOM.ps1` — never hand-edit hashes.

### Adding a dependency

1. Add an entry to `deps.lock.json` under `dependencies[]` with:
   - `name`, `group`, `version`
   - `purl` (package-url format)
   - `license`
   - `type` (library / framework / application)
   - `scope` (required / optional)
   - `coverage` (runtime / development / optional)
   - `source` (human-readable description)
   - `paths` (non-empty array of candidate install locations)

2. Run `.\Build-SBOM.ps1` on a machine where the dependency is
   installed. If the path cannot be resolved, either install the
   dependency or pass `-AllowMissingHashes` (development builds only,
   never for releases).

3. Update `COMPLIANCE.md` runtime component table.

4. Update `SECURITY.md` dependency vulnerability scanning section.

5. Update `INVENTOR_API_GOTCHAS.md` if the dependency is a COM
   component with known quirks.

6. Add a `CHANGELOG.md` entry under the release that ships the change.

### Updating a dependency version

1. Update `deps.lock.json` with the new version and purl.
2. Run `.\Build-SBOM.ps1`.
3. Confirm the SBOM diff shows the expected change.
4. Run `PRJ_SelfTest` on the new host if the change affects Inventor
   compatibility.
5. Add a `CHANGELOG.md` entry.

### What never goes in `deps.lock.json`

- Hashes. They are computed at build time from actual files.
- Transitive dependencies. They belong under `knownUnknowns` with a
  description of why they cannot be enumerated.
- Development-only tools (`a_modDiag`, `a_modExport`). They are
  excluded from `src/` and never ship.

### Known unknowns

`deps.lock.json` has a `knownUnknowns` array. It documents
dependencies that exist but cannot be pinned:

- **transitive** — components pulled in by another dependency
  (e.g., Inventor API transitive COM servers)
- **runtime-resolved** — components located at runtime by ID or
  name (e.g., PDF Translator AddIn GUID)
- **platform-native** — OS-provided components whose versions track
  Windows Update

CISA 2026 accepts an explicit declaration of unknowns as an
alternative to enumeration. Do not remove entries from
`knownUnknowns` unless the underlying uncertainty is resolved.

## Release checklist

Before tagging a release:

- [ ] All tests pass (`RunAllTests`)
- [ ] CHANGELOG.md updated with new version section
- [ ] PRJ_VERSION updated in modConstants
- [ ] README.md version line updated
- [ ] docs/ARCHITECTURE.md version line updated
- [ ] docs/KNOWN_LIMITATIONS.md version line updated
- [ ] docs/ROLLBACK.md version line updated
- [ ] docs/INVENTOR_2027_NOTES.md version line updated
- [ ] docs/INVENTOR_API_GOTCHAS.md version line updated
- [ ] docs/ACCESSIBILITY.md version line updated
- [ ] COMPLIANCE.md version line updated
- [ ] SECURITY.md version line updated
- [ ] ROADMAP.md version line updated
- [ ] TECH_DEBT.md reviewed (resolved items moved, new items added)
- [ ] Build-SBOM.ps1 default $Version updated
- [ ] deps.lock.json version matches PRJ_VERSION
- [ ] SBOM regenerated with new version
- [ ] Build-SBOM.ps1 completed without -AllowMissingHashes
- [ ] Source exported to src/ via Export_All
- [ ] a_modDiag.bas and a_modExport.bas excluded from src/
- [ ] PRODUCTION_BUILD = True in a_modDiag and modTests for release build
- [ ] .ivb saved
- [ ] .ivb copied to releases/BatchPDFPublisher-vX.Y.Z.ivb
- [ ] sbom.cdx.json copied to releases/sbom-vX.Y.Z.cdx.json
- [ ] Git commit
- [ ] Git tag created
- [ ] Backup created via .\scripts\backup-ivb.ps1

## SBOM

The Software Bill of Materials is generated automatically from the
source tree:

    .\Build-SBOM.ps1 -Version X.Y.Z

Format: CycloneDX 1.5 (JSON). Covers CISA 2026 Minimum Elements.

SBOM includes:

- Project metadata and version
- Source tree SHA-256 hash
- SBOM author signature (SHA-256 of JSON content)
- Runtime dependencies (VBA, Inventor, stdole, MSForms) with mandatory
  hashes
- Structured known unknowns
- Coverage and scope per component

Regenerate after every release. Do not edit sbom.cdx.json by hand.

## Versioning

SemVer: MAJOR.MINOR.PATCH.

- MAJOR: breaking change (registry schema, status string contract)
- MINOR: new feature, backward compatible
- PATCH: bug fix

The version string appears in:

- modConstants.PRJ_VERSION
- README.md
- docs/*.md headers
- COMPLIANCE.md
- SECURITY.md
- ROADMAP.md
- deps.lock.json
- Build-SBOM.ps1 default parameter
- CHANGELOG.md section headers

All must be updated together.

### Registry schema versioning

`STATE_SCHEMA_VERSION` (in modConstants.bas) tracks the shape of the
registry tree. When adding, renaming, or removing a registry key:

1. Bump `STATE_SCHEMA_VERSION`.
2. Add a `Case <old-version>` block in `STATE_MigrateSchema` (modState).
3. Write the migration (rename keys via `REG_Get` + `REG_Set` + delete
   old, set defaults for new keys, clear obsolete entries).
4. Log the transition to AUDIT.
5. Document the change in CHANGELOG.md under the release that ships it.

`STATE_MigrateSchema` runs once per session from `STATE_Init`. It
handles four cases (fresh install, same version, older, newer). See
`ARCHITECTURE.md` section 33.5 for the full contract.

Never change a registry key's meaning without bumping the schema
version — that is the only signal a future user has that their
preferences are about to be invalidated.

## Support

Questions, bug reports, feature requests: open an issue or contact
the maintainer directly.

## Maintainer

**Jugoslav Ilić**  
Email: engilic@gmail.com

For questions, bug reports, or contributions, contact the maintainer directly or open a GitHub issue.
