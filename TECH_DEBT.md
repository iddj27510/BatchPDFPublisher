# Technical Debt Register

Prioritized list of known technical debt.

Priority:

- P1: this quarter
- P2: next quarter
- P3: within year

Effort:

- XS: less than 1 hour
- S: 1 to 4 hours
- M: 0.5 to 2 days
- L: 3 to 5 days
- XL: more than 5 days

## Open items

| ID | Description | DIM | Priority | Effort | Status |
|----|-------------|-----|----------|--------|--------|
| TD-01 | modLang 2700 LOC (documented exception) | Arch | P3 | L | Accepted |
| TD-02 | modHelp 1135 LOC (documented exception) | Arch | P3 | L | Accepted |
| TD-03 | Tight coupling: UI forms call logic directly | Arch | P2 | L | Open |
| TD-04 | No modInventorAPI facade | Integ | P2 | L | Open |
| TD-07 | No controller layer | Arch | P2 | L | Open |
| TD-09 | Integration tests for PDF pipeline | Test | P2 | L | Open |
| TD-10 | No CI/CD pipeline | Test | P2 | M | Open |
| TD-13 | No modInventorAPI version check | Integ | P2 | S | Open |
| TD-15 | CKeyPreview requires explicit control whitelist | UI | P3 | S | Open |
| TD-16 | No depth guard on PRJ_CollectByExt recursion | Arch | P3 | S | Open |
| TD-17 | PRJ_PathCombine with UNC bases untested | Test | P3 | XS | Open |
| TD-18 | LOG_SanitizePII masking patterns untested | Test | P3 | XS | Open |

## In progress

None.

## Resolved

| ID | Description | Resolved in | Date |
|----|-------------|-------------|------|
| TD-R01 | PRODUCTION_BUILD = False in release | v9.17.0 | 2026-10-02 |
| TD-R02 | modAI dead code (176 LOC) | v9.16.0 | 2026-10-01 |
| TD-R03 | LOG_GetPath thin wrapper | v9.17.0 | 2026-10-02 |
| TD-R04 | SC_SetSum / SC_SetSum2 thin wrappers | v9.17.0 | 2026-10-02 |
| TD-R05 | PRJ_FSO silent failure on COM error | v9.17.0 | 2026-10-02 |
| TD-R06 | Run_BatchPDFPublisher monolith | v9.17.0 | 2026-10-02 |
| TD-R07 | No pre-flight TranslatorAddIn check | v9.17.0 | 2026-10-02 |
| TD-R08 | Missing SetFocus in frmReport / frmProgress | v9.17.0 | 2026-10-02 |
| TD-R09 | Missing ControlTipText for some controls | v9.17.0 | 2026-10-02 |
| TD-R10 | No Reset confirmation dialog | v9.17.0 | 2026-10-02 |
| TD-R11 | No accelerators for checkboxes | v9.17.0 | 2026-10-02 |
| TD-R12 | Test suite 5 tests only | v9.17.0 | 2026-10-02 |
| TD-R13 | Shell call sanitization in frmReport | v9.18.0 | 2026-10-02 |
| TD-R14 | Audit markers for destructive operations | v9.18.0 | 2026-10-02 |
| TD-R15 | High contrast mode detection | v9.18.0 | 2026-10-02 |
| TD-R16 | frmScan reentrancy during DoEvents pump | v9.18.1 | 2026-10-02 |
| TD-R17 | X button during scan/batch left loop headless | v9.18.1 | 2026-10-02 |
| TD-R18 | KeyPreview = True does not exist in VBA MSForms | v9.18.1 | 2026-10-02 |
| TD-R19 | modValidate duplicated SC_* helpers | v9.18.1 | 2026-10-02 |
| TD-R20 | frmReport reentrancy during RetryClick DoEvents | v9.18.1 | 2026-10-02 |
| TD-R21 | Test count mismatch across docs (11 vs 12) | v9.18.2 | 2026-10-02 |
| TD-R22 | PRJ_PathCombine, RP_BuildSummaryText, STATE_ValidateConfig untested | v9.18.2 | 2026-10-02 |
| TD-R23 | STATE_ValidateEnum substring bug | v9.19.4 | 2026-10-03 |
| TD-R24 | Report file separator inconsistency (· vs vbTab) | v9.19.4 | 2026-10-03 |
| TD-R25 | Host version guard used wrong major-version numbers | v9.25.42 | 2026-10-04 |
| TD-R26 | Checkbox first-click quirk + programmatic-Click side effect | v9.25.42 | 2026-10-04 |
| TD-R27 | frmProgress / frmReport MouseMove theme flicker | v9.25.42 | 2026-10-04 |
| TD-R28 | Ctrl+Z in txtPath leaves validation label stale | v9.25.42 | 2026-10-04 |

## Won't fix

| ID | Description | Reason |
|----|-------------|--------|
| TD-W01 | DPI awareness above 150 percent | MSForms limitation |
| TD-W02 | Focus indicator on Label buttons | Would require CommandButton redesign |
| TD-W03 | Screen reader AccessibleName | MSForms limitation |
| TD-W04 | RTL languages | Out of scope; no demand |
| TD-W05 | PDF/A compliance | Translator AddIn does not expose option |
| TD-W06 | Touch and pen input | MSForms label-based buttons do not support |
| TD-W07 | CJK font rendering tuning | Host locale handles it |
| TD-W08 | Cancel a single IDW export mid-operation | Inventor API has no cancellation hook |
| TD-W09 | Ctrl+Z auto-refresh without any user interaction | Inventor VBA does not support AddressOf (error 35010); no Win32 timer, no Application.OnTime, no MSForms Timer control. Workaround: RevalidatePathIfChanged on UserForm_MouseMove. See KNOWN_LIMITATIONS.md. |
| TD-W10 | Screen reader announcement of label toggle state | Labels have no .Value / checked property in MSForms. Alt+ accelerators are the keyboard affordance. |

## Notes

### TD-01 / TD-02: modLang and modHelp

Largest single source of module LOC. Both are documented as justified
exceptions (single responsibility, data-only, would trigger
over-modularization if split). Accepted as-is.

### TD-03 / TD-07: Coupling and controller layer

Forms currently call PRJ_* functions directly. A controller layer
would decouple UI from logic and improve testability. Tracked for
v9.20.0.

### TD-04: modInventorAPI facade

ThisApplication is called from 8 modules. A facade would isolate the
API surface and enable version checks and mocking.

### TD-09: Integration tests

The test fixtures T01-T11 are in place but not wired into automated
tests. Wiring requires either a headless Inventor run or a fixture-
based PDF status check. Neither is feasible in VBA.

### TD-15: CKeyPreview control whitelist

`CKeyPreview.Initialize` subscribes to `KeyDown` on a fixed set of
control types (TextBox, ComboBox, ListBox, CheckBox, OptionButton,
CommandButton, ToggleButton, ScrollBar, SpinButton, MultiPage,
TabStrip). A custom control that exposes `KeyDown` but is not on
this list would not participate in form-level ESC handling.

Fix: expose the whitelist as a Public array that can be extended
without editing CKeyPreview internals.

### TD-16: Recursion depth in PRJ_CollectByExt

The recursive folder scan has no explicit depth guard. Recursion
depth is bounded by MAX_PATH (260 chars), which caps practical
depth at well under 100 levels. VBA stack overflow is not catchable
by `On Error GoTo`. Adding an explicit depth guard would cost 5 LOC
and turn an uncatchable crash into a loggable warning.

### TD-17 / TD-18: Untested invariants

`PRJ_PathCombine` with UNC bases and `LOG_SanitizePII` masking
patterns are load-bearing but not covered by the headless test
suite. Both are documented in `ARCHITECTURE.md` section 32.4.

### TD-R26: Checkbox → label toggle refactor (v9.25.42)

Two MSForms quirks stacked on the same control:

1. First-click after programmatic `.Value` assignment did not fire
   the auto-toggle.
2. The programmatic `.Value` assignment itself fired `Click` as a
   side effect, flipping state on every form open.

Resolved by removing CheckBoxes entirely and using single Labels
with an Active/Inactive visual state. Labels have no `.Value` and
no auto-toggle, so neither quirk applies. Workaround stack
(`m_bInitializing` guard, debounce gate, flip + force-sync) gone.

### TD-R28: Ctrl+Z validation refresh (v9.25.42)

MSForms consumes Ctrl+Z inside the Windows Edit control before VBA
sees it. The `Change` event does not fire. Workaround:
`RevalidatePathIfChanged` in frmScan runs from `UserForm_MouseMove`
and re-runs validation when the text differs from `g_sLastFolder`.
Residual limitation (W09) documented in KNOWN_LIMITATIONS.md.

## Adding new items

When you discover debt:

1. Add a row to "Open items" with the next available ID.
2. Assign DIM, priority, and effort.
3. If it's small enough and you fix it in the same session, move it
   directly to "Resolved".