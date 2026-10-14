# KNOWN_LIMITATIONS.md - VBA-DR-01 Batch PDF Publisher

**Version:** 9.25.42
**Last updated:** 2026-10-04

Documented limitations with workarounds. Not bugs - deliberate scope
boundaries and platform constraints.

---

## 
Planned improvements (v10.0.0)
- frmReport ListBox replaced with Label-grid (eliminates IntegralHeight bug)
- MouseScroll hook replaced with native subclassing
- modListBoxFit removed (~1600 LOC reduction)

Application Limitations

### IDW Format Only

Only Autodesk Inventor drawings (.idw) are supported. DWG files are
not processed (different API, different architecture).

**Workaround:** Use Inventor to convert DWG to IDW first.

### Single Source Folder

Batch is limited to one source folder per run. Multiple folders
require multiple runs.

**Workaround:** Run the tool once per folder.

### No Cloud Sync During Export

If source folder is on network share or cloud sync (OneDrive,
SharePoint), files may be locked during export.

**Workaround:** Pause cloud sync during batch, or use local copy.

### Single-Instance Lock

Only one Batch PDF Publisher instance per Inventor session. Second
instance shows "already running" message.

**Workaround:** Complete current batch before starting new one.

### Cancel at File Boundary

Cancellation (ESC or Cancel button) takes effect at the next file
boundary, not immediately. A single IDW export cannot be cancelled
mid-operation.

**Reason:** VBA is single-threaded. The DoEvents loop checks
g_bCancel between files.

---

## Platform Constraints (VBA Runtime)

The following limitations are consequences of the VBA runtime, not
architectural mistakes.

### Single-threaded Runtime

VBA has no background thread support. All operations run on the main
UI thread. For UI responsiveness:

- DoEvents via UI_SleepResponsive (modUIHelpers)
- Throttled progress updates via SC_TickProgress (modValidate)
- PDF_ThrottledDoEvents in modPDFProgress

**Not a substitute for:** real async/await or Task.Run.

**Workaround:** Batch can be cancelled via g_bCancel at file
boundaries (ESC button). Cancelling a single IDW export mid-operation
is not possible.

### Reentrancy via DoEvents

VBA does not have threads, but `DoEvents` pumps the Windows message
queue. During a `DoEvents` call from inside `SC_Validate` or
`PRJ_RunScanSession`, the user can invoke another event handler
while the first one is still on the stack. This is the VBA
equivalent of a race condition.

The project uses three state flags to prevent logic errors from
reentrancy:

- `g_bInProcessLoop` — set during `SC_Validate`; checked at entry
  by every frmScan event handler that can mutate batch-relevant
  state (Browse, Theme, Reset, toggle, txtPath).
- `g_bIsProcessing` — set during `RetryClick`; checked at entry by
  every frmReport handler that reads `g_colFiles` (filter, list,
  copy, export).
- `g_bCancel` — set by Cancel button, ESC key, X button, and
  `UserForm_QueryClose`; polled by the batch loop at each file
  boundary.

See `ARCHITECTURE.md` section 31 for the full concurrency model.

**Cancel semantics:** cancellation takes effect at the next file
boundary. A single IDW export cannot be cancelled mid-operation,
because the Inventor API does not expose a cancellation hook.

**X button during batch:** clicking X on `frmProgress` converts to
a cancel request rather than closing the form. This prevents the
batch from continuing headless.

**X button during scan:** clicking X on `frmScan` during a scan
converts to a cancel request. When idle, X closes the form
normally.

### Form-level keyboard handling

VBA MSForms.UserForm does not have a `KeyPreview` property. That is
a Windows Forms (.NET) feature; FM20.DLL does not implement it.
There is no form-level `KeyDown` event that fires when a child
control has focus.

To get ESC handling regardless of focus, the project uses a
`CKeyPreview` class module that subscribes to `KeyDown` on every
focusable control and forwards events to a form-level callback
(`OnFormKeyDown`). Setting `KeyCode = 0` in the callback consumes
the key.

Consequence: the hook is scoped to known focusable control types
(TextBox, ComboBox, ListBox, CheckBox, OptionButton, CommandButton,
ToggleButton, ScrollBar, SpinButton, MultiPage, TabStrip). A custom
control type added in the future that does not appear on this list
would not participate in the form-level ESC handling.

See `ARCHITECTURE.md` section 31.6.

### No `AddressOf` operator

Inventor VBA does not support the `AddressOf` operator. Any attempt
to pass a VBA procedure pointer to a Win32 API raises runtime error
35010 ("Unexpected error") at the call site, even though the code
compiles.

Consequences:

- **No Win32 `SetTimer`.** Background polling with a callback is not
  possible.
- **No window subclassing.** `SetWindowLong(GWL_WNDPROC)` requires a
  callback for the new window procedure.
- **No Windows hooks.** `SetWindowsHookEx` requires a callback.
- **No enumeration callbacks.** `EnumWindows`, `EnumChildWindows`,
  `EnumFontFamilies`, and similar APIs are unusable.

This is a documented limitation of the Inventor VBA host — the
`AddressOf` operator exists in the VBA language reference but is not
exposed to the host's runtime. Excel and Word VBA both support it;
Inventor does not.

**Practical impact on this project:** the Ctrl+Z undo refresh in
`txtPath` cannot use a timer-based solution (see "Ctrl+Z in txtPath"
below).

### No Database Support

VBA can use ADO but this project does not use a database. All
preferences are in registry (HKCU). No connection strings, no DB
credentials.

### No .NET Serialization

VBA does not support BinaryFormatter, System.Text.Json, or
DataContractSerializer. Serialization is done via:

- ADODB.Stream with UTF-8 (report export)
- SaveSetting / GetSetting (registry)
- Native Print # (log file)

### No COM Release API

VBA does not support Marshal.ReleaseComObject. COM objects are
released via Set oX = Nothing idiom. Less precise than Marshal
release but is the only way in VBA.

**Consequence:** Inventor COM objects may stay in memory longer than
ideal. Mitigation: PDF_CleanupAfterExport always calls
Set oDoc = Nothing in the Cleanup: label, even on error path.

### No Async Testing

VBA test suite (modTests) covers synchronous logic. No tests for
"background" behavior because background does not exist. Tests run
in Immediate Window or via RunAllTests.

### No remote log sink

VBA has no async network I/O. Shipping log records to a remote
collector (syslog, Logstash, SIEM) would require either a blocking
HTTP call on the UI thread — unacceptable — or a custom background
mechanism that VBA does not support.

**Workaround:** Parse `%TEMP%\VBADR\BatchPDF_Log_*.txt` and
`BatchPDF_Audit_*.txt` after session end, or use a file-based log
forwarder (e.g., NXLog with a file input source) configured outside
the add-in.

See `ARCHITECTURE.md` section 35.8.

### No log querying or indexing

Logs are plain files in `%TEMP%\VBADR\`. No search, no filter, no
index. Use `findstr` or `Select-String` for ad-hoc queries.

### No structured metrics

`PDF_LogTiming` emits per-phase timing as log entries, but there is
no counter, histogram, or summary. This is a diagnostics mechanism,
not a metrics system. See `ARCHITECTURE.md` section 27.12.

---

## Accessibility Limitations (MSForms)

Target: WCAG 2.2 AA within MSForms constraints.

### What passes

- Color contrast in both themes (light and dark) — verified to >= 4.5:1
- Keyboard navigation: Alt+ accelerators on all buttons
- ESC handler on every form (via CKeyPreview hook)
- **Focus ring on `txtPath` and `txtSearch`** — `BorderColor` switches
  from `CLR_SEP` to `CLR_ACCENT` on `Enter`, back on `Exit`
- Logical tab order (left-to-right, top-to-bottom)
- ControlTipText on all controls, including those without visible text
- Error messages related to relevant input field (txtPath)
- Reset confirmation dialog for destructive settings action
- High contrast mode detection — skips custom palette when
  Windows HC is active

### What does not pass (platform limitations)

- **Focus indicator on Label-based buttons** — MSForms Label does not
  support custom focus visuals. Workaround: none available.
- **Screen reader AccessibleName / AccessibleDescription** — MSForms
  does not expose these properties. Workaround: ControlTipText is the
  closest equivalent.
- **DPI scaling above 150%** — MSForms does not have `AutoScaleMode`
  (that is a WinForms property). Windows scales the whole window, but
  control layout may not be optimal above 150%.
- **Live region announcements** — dynamic status updates (lblSum
  changes) are not announced to screen readers. MSForms limitation.
- **Scan option toggles announced as static text** — the v9.25.42 scan
  option toggles (Subfolders / Force / Audit) are Labels, not
  CheckBoxes. Labels do not have a `.Value` / checked state. Screen
  readers announce them as static text, not as toggles. Keyboard
  affordance is provided via Alt+S / Alt+F / Alt+A accelerators.

### Keyboard shortcuts

Full shortcut map in `COMPLIANCE.md` and `docs/ARCHITECTURE.md`
section 28.10.

Full WCAG status in COMPLIANCE.md.

### Comprehensive accessibility documentation

Full WCAG 2.2 per-criterion mapping, EN 301 549 V4.1.1 detail,
screen reader behaviour, DOJ rule applicability, and the complete
accepted-limitation list are in `docs/ACCESSIBILITY.md`. This
section is a summary.

### TextBox focus indicator

`txtPath` (frmScan) and `txtSearch` (frmReport) render a visible
focus ring: `BorderColor` switches from `CLR_SEP` to `CLR_ACCENT`
on `Enter`, back on `Exit`. This satisfies WCAG 2.2 SC 2.4.7 for
those two inputs.

Label-based buttons (all the primary actions) cannot receive focus
in MSForms, so no focus indicator is possible for them. Alt+
accelerators are the primary keyboard affordance for those buttons.

### Ctrl+Z in txtPath — validation label staleness

**Symptom:** The user clears a character in the source-folder
textbox, then presses Ctrl+Z to restore it. The text is restored,
but `lblSum` and `lblSum2` remain stale ("No folder selected" or
similar) until the user moves the mouse.

**Cause:** Ctrl+Z in an MSForms TextBox is handled internally by
the Windows Edit control before VBA sees it. The VBA `Change`
event does not fire for Ctrl+Z on any host.

**Attempted workarounds (all rejected):**

- `txtPath_KeyUp` handler: reading `.Text` during KeyUp resets the
  Edit control's undo buffer, which breaks Ctrl+Z entirely.
- `txtPath_MouseDown` handler: same — reading `.Text` kills undo.
- `UserForm_Click` handler: same risk.
- Win32 `SetTimer` + `AddressOf` callback: Inventor VBA does not
  support the `AddressOf` operator (runtime error 35010 at the
  call site).
- Subclassing the Edit control (`GWL_WNDPROC` + `WM_UNDO` hook):
  same `AddressOf` dependency, not viable.
- `Application.OnTime`: Excel/Word only; Inventor Application does
  not expose it.
- MSForms `Timer` control: does not exist in MSForms 2.0.

**Mitigation:** `RevalidatePathIfChanged` runs from
`UserForm_MouseMove`. Any mouse movement over the form (including
a 1px move) refreshes both labels. Pressing Enter in the textbox
runs the full validation immediately via `txtPath_KeyDown`. In
practice, the user always moves the mouse or presses Enter right
after Ctrl+Z.

**Status:** Platform limitation, not a bug. Would require a .NET
add-in to solve (`System.Windows.Forms.Timer`).

**Cross-reference:** `ARCHITECTURE.md` §31 (Concurrency),
`TECH_DEBT.md` won't-fix list (TD-W09).

### Accessibility audit

`PRJ_CheckAccessibility` (in `modSelfTest`) walks every form's
controls and reports interactive controls missing `ControlTipText`
and interactive controls with `TabStop = False`. It is advisory —
a missing tooltip does not fail `PRJ_SelfTest`, but should be fixed
before release. Call from Immediate Window:

    PRJ_CheckAccessibility

Output is written to the Immediate Window.

---

## Network and File System

### Network Timeouts Not Configurable

When reading IDW files from slow network shares, individual file
operations use OS-level timeouts. Project does not have configurable
network timeouts.

**Workaround:** If network is slow, batch will take longer. Cancel
via ESC if needed.

### Consecutive Network Failures Abort Batch

After 5 consecutive network failures, batch aborts with
msg.batch.abort.network reason. This is a safety feature to prevent
hours of retries on a dead network.

**Workaround:** Fix network issue, then retry. Already-processed files
show as "Skipped (Exists)".

### Path Length Limit

Maximum path length is 260 characters (MAX_PATH_LEN). Longer paths
are rejected during validation.

**Workaround:** Enable Windows Long Paths (registry + group policy),
but the project still enforces 260 to remain portable.

---

## Localization

### 13 Languages Supported

en, de, fr, it, es, pt, ru, pl, cs, ja, ko, zh, zht.

### Missing: RTL Languages

Arabic, Hebrew, Persian not supported. MSForms RTL support is limited.

### Missing: Additional LTR Languages

Polish, Czech, etc. are covered. Missing: Swedish, Norwegian, Finnish,
Greek, Turkish, Dutch, etc.

**Workaround:** Add new language by creating L_FillXX procedure in
modLang and adding to LANG_Init dispatch.

---

## Testing

### 43 Unit Tests Only

modTests covers pure functions, path handling, cache, formatters,
status classification, localization composition, config validation,
and error classification. No integration tests against a real
Inventor instance running in CI (VBA does not run headless).

No mocking framework exists for VBA. Tests that need real COM
objects (like PDF export) cannot run in isolation. See
`ARCHITECTURE.md` section 32 for the testability strategy.

### Manual Acceptance Tests

T01_HappyPath through T11_DWG are manual test folders with real
Inventor files (~107 MB total). Not tracked in Git.

**Workaround:** Copy from internal file server, or use your own IDW
files.

### Untested Invariants

The following invariants are load-bearing but not covered by the
headless test suite:

- `LOG_SanitizePII` masking patterns — depends on
  `Environ("USERPROFILE")`, hard to make deterministic.
- `CACHE_EvictIfNeeded` FIFO behaviour at scale — would require
  inserting 10,000+ entries; slow and produces no CI value.
- `PRJ_PathCombine` with UNC bases — only simple drive paths are
  tested.
- `CKeyPreview` dispatch chain — requires a live form.
- `STATE_MigrateSchema` — requires real registry writes.
- `PRJ_CheckHostVersion` — requires a specific Inventor version.
- `PRJ_CheckAccessibility` — advisory, not a pass/fail test.
- All UI-level behaviour (hover, theme, reentrancy guards, label
  toggle layout, `LayoutToggleLabels` auto-size reflow) — requires
  a live form.

Manual verification steps for UI-level concerns are documented in
`ARCHITECTURE.md` section 31.

---

## Distribution

### No Code Signing

The add-in is not digitally signed. Windows SmartScreen may block
installation on some configurations.

**Workaround:** Add Windows Defender exclusion, or use `Unblock-File`
on the downloaded file.

### No Automated Installer

Installation requires manual VBE import or manual .ivb copy.

**Workaround:** See README.md Installation section.

---

*End of KNOWN_LIMITATIONS.md*