## INVENTOR_API_GOTCHAS.md — kompletan

```markdown
# Inventor API Gotchas

Version: 9.25.42
Host: Autodesk Inventor 2027, VBA 7.1

Compiled from v8.00 internal technical notes (TN-01..TN-26), v9.x
development experience, and integration testing.

Corrections in this revision:
- TN-24 Keyboard: prior revision advised `Me.KeyPreview = True` for
  form-level ESC handling. This is a Windows Forms (.NET) feature
  that does NOT exist in VBA MSForms (FM20.DLL). Attempting to set
  it causes a compile error. The project uses a `CKeyPreview` class
  module instead — see the corrected entry and
  `ARCHITECTURE.md` section 31.6.
- TN-20: prior revision described a `MAX_FOLDER_DEPTH = 32` guard
  that does not exist in the code. Corrected to describe the actual
  state (bounded by OS limits, no explicit guard).
- TN-23: prior revision named `GetTickCount64` / `GetTickCount` as
  the timing API. The project uses `QueryPerformanceCounter` /
  `QueryPerformanceFrequency`. Corrected.
- TN-27 (new): Ctrl+Z in MSForms TextBox does not raise Change.
  See the new section at the end of this document.

## PDF export

### TN-01 — `kPrintAllSheets` enum does not exist
The Inventor API does not expose a `kPrintAllSheets` enum. The correct
value for "all sheets" in `Sheet_Range` is `73730` (Long). Passing a
named constant will cause compile failure.

### TN-02 — `NameValueMap` uses `.Add`, not `.Value()` on first call
`NameValueMap` has `Add` and `Value`. On the first write for a given key,
use `.Add(key, value)`. To overwrite, use `.Value(key) = value`. Calling
`.Value(key) = x` for a key that does not exist may silently fail.

Required keys for PDF export:
- `Publish_All_Sheets` = True
- `Launch_Viewer` = False
- `DeferUpdates` = True (optional, speeds up open)
- `SkipAllUnresolvedFiles` = True (on `OpenWithOptions` options map)

### TN-10 — `OpenWithOptions` third argument
The third argument of `ThisApplication.Documents.OpenWithOptions` must
be `False` (do not force open as visible). Passing `True` opens the
document in the UI, which is not desired for batch processing.

### TN-24 — `DoEvents` placement
`DoEvents` inside the per-file loop is safe only **before**
`SaveCopyAs`. After `SaveCopyAs`, the document COM reference may be
invalidated by the message pump. In `PRJ_CreatePDF`, all `DoEvents`
calls happen via `PRJ_UpdateProgress` which is called at defined points.

## Open-drawing rule

### Own what you open
- Before calling `OpenWithOptions`, check `Documents.ItemByName(sIDW)`
  to see if the document is already open.
- Track `bWeOpened` = True only if **this add-in** opened the document.
- On cleanup, close only if `bWeOpened` is True.
- A document that was open before the add-in ran stays open after.

## COM and MSForms

### TN-09 — `ThisApplication.VBE`, not `Application.VBE`
`Application.VBE` raises Error 424 in Inventor VBA. Use
`ThisApplication.VBE`. For the diagnostic dump module, this is the
entry point for enumerating components.

### TN-13 — `VBA.UserForms.Add("frmNaziv")`
To show a form instance from outside its own code, use
`VBA.UserForms.Add("frmNaziv")`. Direct `New frmNaziv` may fail to
correctly initialize.

### COM object release
Every `CreateObject` or `Set x = SomeObject.Child` should be matched
with `Set x = Nothing` in all exit paths, including error handlers.
This is critical for long-running batches.

## Forms and layout

### TN-06 — Form geometry
- Set `Width`, `Height`, `StartUpPosition` only in `UserForm_Activate`,
  never in `UserForm_Initialize`.
- `Me.Caption` is set in `UserForm_Initialize`.
- In Designer, caption may show "ghost text" — ignore it.

### TN-07 — Label-as-button
- `SpecialEffect = 0` (Flat)
- `BackStyle = 1` (Opaque)
- `BorderStyle = 0` (None)
- `TextAlign = 2` (Center)
- `TextAlign = 0` raises Error 380.

### TN-15 — `AutoSize = True` on controls
Autosize is discouraged on labels used as button frames because long
translations break the layout. Where dynamic sizing is needed (e.g.
`lblSearch`, or the v9.25.42 scan option toggles), use runtime
recalculation with a right-edge anchor.

**Order-critical quirk:** MSForms ignores `.Width` / `.Height`
assignments while `AutoSize` is True. To set an explicit width on an
auto-sizing label, first set `AutoSize = False`, then assign
`.Width` / `.Height`. The v9.25.42 `LayoutToggleLabels` procedure
in frmScan does this explicitly and documents why.

### TN-16 — `Font.Bold` and `Font.Weight`
Set **both** `Font.Bold = True` and `Font.Weight = 700`. Setting only
one may not have the expected visual effect across themes.

### TN-17 — Progress bar painting
`progress-bar` = track (BackColor = CLR_SEP).
`progress-fill` = fill (BackColor = CLR_ACCENT, Width scaled).

### TN-24 — Form-level keyboard handling (corrected)
VBA MSForms.UserForm does **not** have a `KeyPreview` property.
That is a Windows Forms (.NET) feature; FM20.DLL does not implement
it. There is no form-level `KeyDown` event that fires when a child
control has focus.

**WRONG (causes compile error):**
    Me.KeyPreview = True

**CORRECT:** use the `CKeyPreview` class module that subscribes to
`KeyDown` on every focusable control and forwards events to a
form-level callback:

    Set m_oKeyPreview = New CKeyPreview
    m_oKeyPreview.Initialize Me, "OnFormKeyDown"

The callback:

    Public Sub OnFormKeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
        If KeyCode = 27 Then
            ' handle ESC
            KeyCode = 0    ' consume
        End If
    End Sub

Setting `KeyCode = 0` consumes the key, matching `Handled = true` in
.NET. Release the hook in `UserForm_Terminate`:

    If Not m_oKeyPreview Is Nothing Then
        m_oKeyPreview.Terminate
        Set m_oKeyPreview = Nothing
    End If

See `CKeyPreview.cls`, `CKeyPreviewHook.cls`, and
`ARCHITECTURE.md` section 31.6.

## Theme detection

### TN-14 — ThemeManager API
- `ThisApplication.ThemeManager.ActiveTheme.Name` — modern Inventor.
- `AppearanceManager` does **not** exist — Error 438.
- Fallback: `ThisApplication.ColorSchemes.ActiveColorScheme.Name`.
- Empty name is treated as Light.

### Windows dark mode
Read `HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme`:
- `0` = Dark
- `1` = Light
Use `WScript.Shell.RegRead`.

## File collection

### TN-20 — `Dir()` is not reentrant
Recursive folder scanning must not use `Dir()` recursively. Use
`Scripting.FileSystemObject` and recursion with a bounded call stack.

Our implementation: `PRJ_CollectByExt` (modSummary) uses
`Scripting.FileSystemObject` and recursion into `oFld.SubFolders`.
There is **no explicit `MAX_FOLDER_DEPTH` guard** — recursion depth
is bounded by Windows path length limits (MAX_PATH = 260 chars),
which caps practical depth to well under 100 levels. VBA stack
overflow is **not** catchable by `On Error GoTo`, so if a folder
tree ever exceeds the VBA stack limit, the host crashes. This is
accepted as a practical non-concern for the target use case
(engineering project folders, typically <10 levels deep).

### TN-21 — OldVersions skip
Skip `OldVersions` at **folder** level, not file by file. Check
`LCase(sub.Name) <> "oldversions"` before recursing.

## Timing

### TN-23 — Midnight-safe elapsed (corrected)
`Timer()` returns seconds since midnight local time. If a batch
crosses midnight, `Timer() - tStart` is negative. Naive fix is
`If d < 0 Then d = d + 86400`.

Our implementation: `TM_Now()` / `TM_Elapsed()` in `modTime`, based
on **`QueryPerformanceCounter` / `QueryPerformanceFrequency`** from
kernel32. Monotonic, DST-safe, never wraps in practice. Falls back
to `Timer()` if QPC is unavailable (rare — QPC is present on all
supported Windows versions since XP).

### DST transitions
`Timer()` jumps backward on DST fall-back. `QueryPerformanceCounter`
does not. Any timing that spans more than a few minutes must use the
monotonic timer.

## Clipboard

### TN-26 — HTMLfile primary, CLSID fallback
Two clipboard strategies:
1. `CreateObject("htmlfile").ParentWindow.ClipboardData.SetData "Text", s`
   — fast, UTF-8 safe on modern Windows.
2. `CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")` — OLE
   DataObject, `SetText` + `PutInClipboard` — fallback for older hosts.

Both must be wrapped in `On Error` blocks. If both fail, fall back
to `cmd.exe /c clip < tempfile` with sandbox path (never user path).

## StatusBar

### TN-18 — StatusBarText
While processing:
`ThisApplication.StatusBarText = "PDF Export [i/total]  filename"`.

After processing:
`ThisApplication.StatusBarText = "PDF Export finished"` or
`"PDF Export finished with errors"`.

## Diagnostics

### Diagnostic dump
`a_modDiag` provides three modes: `DiagDump_Code`, `DiagDump_Props`,
`DiagDump_Full`. It parses the `.frm` header block to read properties
that the VBE host does not expose (KeyPreview, ControlBox, etc.).

### Extension
For long-running diagnostic output, `DoEvents` is called every 500
lines. This keeps the host responsive but adds overhead; disable for
production use with `#Const PRODUCTION_BUILD`.

## Project-specific patterns

### Full Jitter retry
    delay = Int(Rnd() * Min(cap, base * 2^attempt))
    If delay < 50 Then delay = 50

This is the AWS-recommended pattern. Deterministic backoff causes
thundering herd. See `modRetry`.

### Atomic PDF write
1. Export to temp file `_tmp_<PID>_<unique>.pdf` in sandbox.
2. Validate PDF header (`%PDF-`).
3. If target exists, move target to `.bak`.
4. Move temp to target.
5. If move fails, restore `.bak` to target.
6. On success, delete `.bak`.
7. On mismatch, rollback: delete new target, restore `.bak`.

See `PDF_SafeReplace`, `PDF_CommitReplace`, `PDF_RollbackReplace`
in `modPDFFileIO`.

### Cache eviction
`Scripting.Dictionary` — no built-in LRU. FIFO eviction via
`Keys()` array (insertion order). Cache bounded at
`CACHE_MAX_ENTRIES = 10000`, evict `CACHE_EVICT_BATCH = 2000` when
exceeded. See `CACHE_EvictIfNeeded` in `modCache`.

### Path sanitization
Single function `SN_SanitizePath` handles:
- outer quote stripping
- illegal char removal (control, `"*?<>|`)
- duplicate backslash collapse (preserving UNC prefix)
- parent segment resolution (`..`)
- UNC share guard (reject if traversal escapes share)
- length check

See `modPath`.

## Undo (Ctrl+Z) in MSForms TextBox

### TN-27 — Ctrl+Z bypasses the Change event and resets on .Text read

Ctrl+Z in an MSForms TextBox is handled by the underlying Windows
Edit control. Three consequences:

1. **The VBA `Change` event does not fire.** Any validation or
   side effect that depends on `Change` will not run after an undo.
2. **Reading `.Text` during `KeyUp` or `MouseDown` resets the undo
   buffer.** Adding a `txtPath_KeyUp` handler that reads `.Text`
   breaks Ctrl+Z entirely — after the first keystroke the undo
   history is gone.
3. **A `UserForm_Click` handler that reads `.Text` has the same
   risk** on some hosts.

The only reliable workaround in Inventor VBA is a **passive
polling-free detector**: compare the current text against the last
known value inside an event that does not read the Edit control
during its own message pump. In this project, that is
`UserForm_MouseMove`:

    Private Sub UserForm_MouseMove(...)
        ...
        RevalidatePathIfChanged
    End Sub

`RevalidatePathIfChanged` reads `.Text` once per mouse move and
re-runs validation only when the text differs from `g_sLastFolder`.

**Attempted alternatives (all rejected):**

| Approach | Failure mode |
|---|---|
| `txtPath_KeyUp` handler | Reads `.Text` → resets undo buffer |
| `txtPath_MouseDown` handler | Same |
| `UserForm_Click` handler | Same on some hosts |
| Win32 `SetTimer` + callback | Inventor VBA: `AddressOf` raises error 35010 |
| Subclassing Edit control (WM_UNDO hook) | Same `AddressOf` dependency |
| `Application.OnTime` | Excel/Word only; Inventor Application has no `OnTime` |
| MSForms `Timer` control | Does not exist in MSForms 2.0 |

**Residual limitation:** pure keyboard Ctrl+Z with no subsequent
mouse movement does not refresh labels until the user moves the
mouse. In practice the user always moves the mouse or presses Enter.
See `docs/KNOWN_LIMITATIONS.md` §"Ctrl+Z in txtPath" for the full
analysis.

**Cross-reference:** `modScanPoll` (Win32 SetTimer attempt) was
written and deleted during v9.25.42 development for exactly this
reason. See CHANGELOG.md.

## VBA ceiling

Realistic maximum for VBA in Inventor host: ~8.8 / 10.

Reasons:
- No async/await
- No dependency injection
- No responsive layout
- No unit test framework (we built minimal `modTests`)
- No CI/CD
- No `KeyPreview` (form-level keyboard requires CKeyPreview hook)
- No `AddressOf` (blocks all Win32 timer / hook / subclass patterns)
- COM / MSForms limitations (font rendering, pixel control, theming)

## See also

- `docs/ARCHITECTURE.md` — module map and invariants
- `docs/KNOWN_LIMITATIONS.md` — documented limitations and workarounds
- `docs/ROLLBACK.md` — rollback procedure
- `modState` header — state lifecycle contract