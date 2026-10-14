# Batch PDF Publisher - Installation and Usage Guide

This guide covers everything from "I have never opened the VBA editor"
to "I want to deploy this to 40 workstations via script". Pick the track
that matches you; both end at the same working toolbar button.

Applies to: Autodesk Inventor 2026 and 2027, Windows 10/11, VBA 7.1 host.
The add-in itself is pure VBA (no compiled DLL, no external runtime).

## Before you start (both tracks)

You need three things:

1. Inventor 2026 or 2027 installed and able to open a drawing.
2. The add-in files. Either the packaged `BatchPDFPublisher.ivb`, or the
   `src/` folder from the repository (40 modules, 8 classes, 3 forms).
3. (Optional but strongly recommended) Apprentice Server registered.
   Without it the tool still works, but reading sheet counts from IDW
   files is 10-50x slower because it falls back to the full Inventor
   engine. The in-app dialog and `docs/INVENTOR_2027_NOTES.md` explain
   how to register it; this guide assumes you either did or accepted
   the slower path.

If you downloaded the `.ivb` or the repo zip from the internet, Windows
may block it. Before importing: right-click the file > Properties > if
you see an "Unblock" checkbox at the bottom, tick it > Apply. Skipping
this is the single most common reason "the macro does not appear".

## Track A - Total beginner

Goal: get a clickable button on an Inventor ribbon, in about ten minutes,
without touching the registry or writing code.

### A.1 Open the VBA editor

Start Inventor. In the ribbon go to Tools > VBA Editor (or press Alt+F11).
A separate window opens with a "Project Explorer" panel on the left. If
you do not see that panel, press Ctrl+R inside the VBA window.

In Project Explorer you will see one or more entries like:

  VBAProject (ThisDrawing)
  VBAProject (Standard)
  VBAProject (GlobalMacros)   <- this one may or may not be there yet

### A.2 Choose where the code lives (read this once, it matters)

There are two places VBA code can live, and the difference is whether the
button survives a restart of Inventor.

- Standard: per-session. Code you import here disappears when you close
  Inventor. Good for a one-off test, bad for daily use.
- GlobalMacros: persistent. Code imported here is reloaded automatically
  every time Inventor starts, on this user account. This is what you want.

If `GlobalMacros` is not listed in Project Explorer, that is normal on a
fresh install. Inventor creates it the first time you import into it, or
you can create it via Tools > Macros > (the GlobalMacros project is the
default target). The simplest reliable path: import into GlobalMacros;
if the IDE refuses, import into Standard first, then use the method in
B.1 below to relocate.

### A.3 Import the project

In Project Explorer, right-click GlobalMacros > Import File...

Navigate to the folder holding the exported sources. Select every `.bas`,
`.cls`, and `.frm` file. To save time, in the file dialog press Ctrl+A to
select all, then Open. VBA will ask nothing further; it just loads them.

If you have the single packaged `BatchPDFPublisher.ivb` instead of loose
files: right-click GlobalMacros > Import File... and pick the `.ivb`. One
file, done.

After import, expand GlobalMacos in Project Explorer. You should see
Modules (modEntry, modBatch, modReportList later, ...), Class Modules
(CHook*, CKeyPreview, ...), and Forms (frmScan, frmReport, frmProgress).
If any of those three folders is missing, the import was partial - redo
A.3 selecting all files.

### A.4 Compile before you run

Menu: Debug > Compile VBAProject. It must finish with no error dialog.
If it stops on "Variable not defined", a module did not import (see A.3)
or you mixed two versions of the source. Fix that now; running an
uncompiled project is how people conclude "the tool is broken" when it
is actually half-loaded.

### A.5 Run it once by hand

In Project Explorer double-click `modEntry`, find the `Run_BatchPDFPublisher`
sub, put the cursor inside it, press F5. The scan form appears. Pick a
folder with `.idw` files, leave the defaults, press Run. Watch the
progress form, then the report form. Close it. You now know the tool
works; the button in A.6 just automates this exact step.

(Alternative without opening code: Tools > Macros, select
`Run_BatchPDFPublisher` from the GlobalMacros project, Run. Same result.)

### A.6 Make a toolbar button (the part everyone skips and then regrets)

Two ways. Use the first; keep the second as a fallback if the macro does
not show up in the list.

Method 1 - drag and drop (no code):

1. In Inventor (not the VBA window): Tools > Customize.
2. Go to the Commands tab.
3. In the Categories list on the left, scroll to Macros. If your project
   imported cleanly, the macros from GlobalMacros appear here. Look for
   `Run_BatchPDFPublisher`.
4. In the Commands list on the right, find that macro, click and hold it,
   and drag it onto any ribbon tab or the Quick Access Toolbar. Drop it.
5. Rename it while still in Customize: select the new button, in the
   Modify Selection box type a short name like "Batch PDF". Optional:
   Assign Icon to give it a picture instead of text.
6. Close Customize. The button is now on the ribbon and persists across
   restarts because it points at a GlobalMacros macro.

If `Run_BatchPDFPublisher` is NOT under Macros in step 3: the project is
not in GlobalMacros (it is in Standard, which does not register macros
for the Customize list reliably). Move it per B.1, restart Inventor, and
retry. A second cause is the Unblock step at the top of this guide.

Method 2 - one-time VBA snippet (when you prefer code over clicking):

Open the VBA editor, insert a standard module (Insert > Module), paste:

  Sub InstallBatchPdfButton()
      Dim cb As CommandBar
      On Error Resume Next
      Set cb = ThisApplication.CommandBars("BatchPDFPublisher")
      On Error GoTo 0
      If cb Is Nothing Then Set cb = ThisApplication.CommandBars.Add( _
          "BatchPDFPublisher", msoBarType.msoBarTypeNormal, , True)
      Dim btn As CommandBarButton
      Set btn = cb.Controls.Add(msoControlButton)
      btn.Caption = "Batch PDF"
      btn.OnAction = "Run_BatchPDFPublisher"
      btn.Style = msoButtonCaption
      cb.Visible = True
  End Sub

Run it once (F5 inside the sub). It creates a dedicated toolbar with the
button wired to the global macro. You can drag that toolbar anywhere or
merge the button into an existing ribbon afterwards via Customize. To
remove it later, delete the toolbar in Customize or run a one-liner
`ThisApplication.CommandBars("BatchPDFPublisher").Delete`.

`OnAction = "Run_BatchPDFPublisher"` works without qualification only
because the macro lives in GlobalMacros. If you installed into Standard,
the button breaks on the next restart - which is exactly why A.2 insists
on GlobalMacros.

### A.7 First real run, narrated

Click your new button. The scan form opens.

- The big path box at top: type or browse to the folder holding your
  drawings. The Browse button opens a folder picker.
- Three checkboxes under it: Subfolders (scan recursively, skips
  OldVersions), Force overwrite (recreate PDF even if newer than the IDW),
  Audit (also inspect .ipt/.iam for missing drawings). Defaults are sane;
  leave them on first try.
- Output options: single combined PDF vs individual files per drawing.
- Theme and language selectors: see "Languages and themes" below.
- Press Run. The progress form shows a per-file line and a bar. When it
  finishes, the report form opens automatically.

In the report form:

- The grid lists every processed file with Status, Size, Pages, Modified,
  Path. Status carries a marker (a check, a bang, a tilde, etc.) plus a
  word.
- Click a row: the full path appears in the line under the grid (and as a
  tooltip on hover once the label-grid build lands).
- Double-click a row: opens the PDF if it exists, otherwise opens the IDW
  in Inventor. This is the fastest way to eyeball a result.
- The five filter buttons (All / Errors / Skipped / No .idw / OK) narrow
  the grid. Search box narrows by path substring, live.
- Column headers sort: first click ascending, second descending, third
  clears the sort.
- Copy puts the whole report on the clipboard. Export writes it to a file
  and opens Notepad. Retry re-runs only the errored rows. Back returns to
  the scan form. Cancel closes the report.

That is the entire daily loop. Everything below is reference.

## Track B - Professional / deployment

For people who install this on many machines, script it, or build from
source.

### B.1 Relocate Standard -> GlobalMacros (if A.2 put you in Standard)

In the VBA editor, with the project loaded in Standard: export every
component (File > Export File, or run the project's own `Export_All`
helper if present in modDev), then import into GlobalMacros per A.3, then
remove the Standard copies. Net effect: identical code, persistent home.
Restart Inventor and confirm the macro still appears under Customize >
Macros.

### B.2 Find the GlobalMacros path for scripting

Right-click GlobalMacros in Project Explorer > Properties. The Location
field is the per-user folder (under %APPDATA%\Autodesk\Inventor <ver>\...).
Use that exact path in your deploy script rather than hardcoding it,
because it differs between 2026 and 2027 and between machines. Copy the
`.ivb` (or the loose sources, depending on how you package) into that
folder, or import via the VBE object model if you automate the editor.
For fleet deployment the usual pattern is: drop the `.ivb` to the
GlobalMacros location via SCCM/GPO/logon script on first launch, then let
Inventor load it. No registry edits are required for the VBA path; the
registry-based add-in mechanism is for compiled COM add-ins and is out of
scope for this pure-VBA tool.

### B.3 Build and self-test from source

The repository ships `modSelfTest.bas` and `modTests.bas` (49 tests).
After importing sources, run the self-test entry point from modEntry (or
Tools > Macros, pick the self-test sub). It exercises path handling, PDF
metadata, cache, report formatting, localization and the regression
guards without needing a live drawing. UI-level behavior (hover, theme,
label toggle, report grid) is not covered headlessly - that is the
T01-T11 manual fixture pass documented in `docs/ARCHITECTURE.md` and
`CONTRIBUTING.md`. Run one manual pass per release on your machine.

### B.4 Versioning and rollback

Source of truth for the version string is `PRJ_VERSION` in
`modConstants.bas`; the form caption and the SBOM both read from it. The
repository is tagged per release (e.g. `v9.25.42`). To roll a workstation
back: replace the GlobalMacros `.ivb` with the tagged one and restart
Inventor. `ROLLBACK.md` lists the exact steps and the known-good tags.
Do not edit `PRJ_VERSION` by hand outside a release commit; the SBOM
generator (`Build-SBOM.ps1`) and `deps.lock.json` are aligned to it and a
mismatch trips the compliance check.

### B.5 Compliance and supply chain

`docs/COMPLIANCE.md`, `SECURITY.md`, and the generated `sbom.cdx.json`
(CycloneDX 1.5) cover the CISA 2026 minimum elements. The four declared
dependencies (VBA 7.1, Inventor 2027, stdole 2.0, MSForms 2.0) are
hash-verified at build time by `Build-SBOM.ps1`; transitive Inventor COM
servers and the runtime-resolved PDF Translator AddIn GUID are declared
as known-unknowns per the documented exception in `ARCHITECTURE.md`.
Regenerate the SBOM after any source change that ships.

## Using the report - reference

Columns (fixed order, widths measured once from the unfiltered set so
they never jump when you filter): Status, Size, Modified, Pages, Path.
Path is the elastic column; it takes whatever width is left.

Filters: All, Errors, Skipped, No .idw, OK. They are mutually exclusive;
picking one reloads the grid with that subset. The grid geometry does not
change between filters (that was the whole point of the v9.25.42 fix).

Search: free-text, matches against the relative path, live as you type.
Combined with the active filter (search narrows within the filter).

Sort: click a column header. Ascending, then descending, then unsorted.
The header shows a triangle for the active sort direction.

Keyboard (report form). Note the trap: the Alt numbers for the last two
filters do NOT follow the left-to-right button order. Buttons read
All, Errors, Skipped, No .idw, OK; the shortcuts are Alt+1 All, Alt+2
Errors, Alt+3 Skipped, Alt+4 OK, Alt+5 No .idw. So Alt+4 jumps to the
rightmost button and Alt+5 to the second-from-right.

  Alt+1   filter All
  Alt+2   filter Errors
  Alt+3   filter Skipped
  Alt+4   filter OK          (not "No .idw", despite button position)
  Alt+5   filter No .idw     (not "OK", despite button position)
  Alt+B   Back to scan form
  Alt+C   Copy report to clipboard
  Alt+R   Retry errored rows
  Alt+E   Export report to file
  Alt+Q   Cancel / close report
  Ctrl+F  focus the search box
  Esc     if search has text: clear it; else: close the report
  Up/Down/PageUp/PageDown/Home/End  move the selection (label-grid build)
  Enter   open the selected row (same as double-click)
  Double-click row   open PDF if present, else open IDW in Inventor
  Click row          show full path under the grid

Mouse: wheel scrolls the grid; the thin scrollbar on the right can be
dragged or page-clicked (label-grid build). Hovering a row shows the full
path as a tooltip once the label-grid build is in; on the legacy ListBox
build the path appears in the line under the grid on click.

Summary line: counts of created / overwritten / skipped / errors plus
elapsed time, colored by outcome (green when clean, amber when skipped,
red when errors exist). In audit mode a second line shows model coverage
(with drawing / without drawing) from the cached or live scan.

## Languages and themes

13 UI languages; the product name itself is never translated (it is a
constant, `APP_TITLE`). Switch language from the scan form selector; the
whole UI re-localizes on the next form activation, and the report grid
re-measures its column widths for the new strings (measurement runs after
theme + localization, never before - that ordering is load-bearing).

Four themes: Inventor (brand palette, dark), Light, Dark, System
(follows Windows). Theme is applied by tag dispatch (`FW_ApplyTheme`
reads each control's Tag, not its Name), so adding a control only needs a
correct Tag in the designer. Contrast ratios for the brand palette are
documented in `docs/ACCESSIBILITY.md` (WCAG AA target).

## Troubleshooting

Apprentice Server warning on startup. The tool checks registration and
shows a short guide if missing. Register it per the in-app dialog or
`docs/INVENTOR_2027_NOTES.md`; or accept the slower full-engine path. The
warning can be suppressed via the registry key named in the in-app text,
but suppressing does not make it fast.

Macro does not appear under Customize > Macros. Three causes, in order of
likelihood: (1) the `.ivb`/sources were blocked by Windows - redo the
Unblock step; (2) the code is in Standard, not GlobalMacros - do B.1 and
restart Inventor; (3) the import was partial - Debug > Compile must be
clean and Project Explorer must show Modules + Class Modules + Forms.

"Variable not defined" on Compile. A module failed to import, or two
versions of the same module are both loaded (duplicate names across
Standard and GlobalMacros). Remove duplicates, re-import cleanly, compile.

iLogic toggles off during a batch and you did not ask it to. By design.
Opening/closing documents with SkipAllUnresolvedFiles can crash the .NET
runtime via an Autodesk bug; the add-in disables iLogic for the duration
of the batch and restores it afterward. You do not need to do anything.
If iLogic is still off after the batch ends, the run was interrupted -
re-enable it in Tools > Add-Ins.

32-bit vs 64-bit. The project targets VBA7 and uses PtrSafe declarations
where it calls Win32. It runs in both host bitnesses. If you see
"different declaration" compile errors on a Win32 API call, the host is
older than VBA7 - upgrade Inventor/VBA, do not patch the declare.

High DPI / scaling (125%, 150%). Text measurement is in points via GDI
and DPI-aware; column widths and row heights are computed from the real
font metrics, not hardcoded pixels. If a build ever shows clipped rows or
wrong widths at non-100% scaling, that is a regression in the measurement
path - report it with the scaling percentage and the `[RL]` debug line
from the Immediate window.

Scroll does nothing / last row clipped (legacy ListBox builds). This is
the IntegralHeight partial-row bug that the v10 label-grid migration
removes. On v9.25.42 it is mitigated by snapping list height to whole
rows before freezing; if you still hit it on a specific filter, note the
filter and row count - it is the exact case the new grid eliminates by
construction.

Where are the logs. `%TEMP%\VBADR\BatchPDF_Log_<PID>.txt`, one per
process, 90-day retention. The report's Copy/Export actions do not write
here; this is the structured session log (start, end, per-file outcomes,
warnings). Useful when filing a bug: attach the log and the exported
report.

## Uninstall

Remove the button: Tools > Customize, drag it off the ribbon (or delete
the toolbar if you used method B's snippet). Remove the code: in the VBA
editor, right-click each imported component under GlobalMacros > Remove
File... (answer "No" when asked to export a copy, unless you want one).
Or delete the `.ivb` from the GlobalMacros folder found in B.2 and
restart Inventor. Nothing is written to the registry by the VBA install
path, so there is no registry cleanup step. Preferences the tool stored
under `HKCU\Software\VB and VBA Program Settings\VBADR01\` can be deleted
for a factory reset, but leaving them is harmless.

## File map (what lives where)

Source layout and module responsibilities are authoritative in
`docs/ARCHITECTURE.md`. Quick orientation for this guide: `modEntry`
holds `Run_BatchPDFPublisher` (the button target); `modScan`/`frmScan`
are the input form; `modBatch`/`modPDFExport`/`modPDFFileIO` are the
engine; `modReport*` + `frmReport` are the report UI (the label-grid
migration lives in `modReportList` and, later, `modWheelHook`);
`modLang`/`modTheme`/`modPalette`/`modFont` are localization, color and
text measurement; `modSelfTest`/`modTests` are the headless suite.

## Keyboard reference (repeat, most-used section)

  Alt+1 All | Alt+2 Errors | Alt+3 Skipped | Alt+4 OK | Alt+5 No .idw
  Alt+B Back | Alt+C Copy | Alt+R Retry | Alt+E Export | Alt+Q Cancel
  Ctrl+F focus search | Esc clear-search-then-close
  Click row = full path | Double-click row = open PDF/IDW
  Header click = sort asc / desc / none
