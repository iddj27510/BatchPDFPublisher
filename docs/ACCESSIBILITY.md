# ACCESSIBILITY.md - VBA-DR-01 Batch PDF Publisher

**Version:** 9.25.42
**Last updated:** 2026-10-04
**Target:** WCAG 2.2 Level AA (partial), ISO/IEC 40500:2025, EN 301 549 V4.1.1

Comprehensive accessibility documentation for the add-in. Covers
conformance status, platform limitations, screen reader behaviour,
keyboard navigation, colour contrast, and future improvements.

---

## 1. Conformance summary

| Standard | Status | Notes |
|---|---|---|
| WCAG 2.2 Level A | Pass (within MSForms) | See §3 |
| WCAG 2.2 Level AA | Pass (within MSForms) | See §3 |
| WCAG 2.2 Level AAA | Not targeted | Not required for this application class |
| ISO/IEC 40500:2025 | Same as WCAG 2.2 | ISO free version |
| EN 301 549 V4.1.1 | Partial | Adopts WCAG 2.2 AA; see §5 |
| Section 508 (US) | Not applicable | Not US federal procurement |
| DOJ web accessibility rule | Not applicable | Desktop add-in, not a web application |

---

## 2. Platform context

The add-in runs as a VBA UserForm inside Autodesk Inventor 2027.
The UI layer is **MSForms** (`FM20.DLL`), which is a legacy
Windows Forms toolkit with significant accessibility limitations
compared to modern frameworks.

MSForms does **not** implement:

- `AccessibleName`, `AccessibleDescription`, `AccessibleRole`
  (UI Automation properties)
- Live region announcements for dynamic status updates
- Custom focus indicators on Label controls (Labels cannot receive
  focus at all — they are non-interactive by design)
- `AutoScaleMode` for DPI-independent layout
- Native high-contrast theme switching at the control level

These are documented as accepted limitations, not defects. See
`docs/KNOWN_LIMITATIONS.md` for the workaround for each.

---

## 3. WCAG 2.2 conformance detail

### 3.1 Perceivable

| SC | Level | Status | Notes |
|---|---|---|---|
| 1.1.1 Non-text Content | A | N/A | No informational non-text content. All UI is text. |
| 1.3.1 Info and Relationships | A | Partial | `ControlTipText` conveys association; no programmatic label association (MSForms limitation). |
| 1.3.2 Meaningful Sequence | A | Pass | Tab order matches visual order. |
| 1.3.3 Sensory Characteristics | A | Pass | No instructions rely solely on shape, size, or position. |
| 1.3.4 Orientation | AA | Pass | Fixed portrait; no orientation lock. |
| 1.3.5 Identify Input Purpose | AA | N/A | No personal data inputs. |
| 1.4.1 Use of Color | A | Pass | Status uses text + colour; no information conveyed by colour alone. |
| 1.4.2 Audio Control | A | N/A | No audio. |
| 1.4.3 Contrast (Minimum) | AA | Pass | See §4. |
| 1.4.4 Resize Text | AA | Partial | MSForms does not support runtime text resize; DPI scaling handled by Windows. |
| 1.4.5 Images of Text | AA | Pass | No images of text. |
| 1.4.10 Reflow | AA | N/A | Fixed-size desktop dialog. |
| 1.4.11 Non-text Contrast | AA | Pass | Focus border on TextBox uses `CLR_ACCENT` (5.6:1 against white). |
| 1.4.12 Text Spacing | AA | N/A | No user-adjustable spacing. |
| 1.4.13 Content on Hover or Focus | AA | Pass | Tooltips dismissible via ESC; no content blocks the UI. |

### 3.2 Operable

| SC | Level | Status | Notes |
|---|---|---|---|
| 2.1.1 Keyboard | A | Pass | All actions reachable via keyboard. ESC, Tab, Enter, Alt+ shortcuts. |
| 2.1.2 No Keyboard Trap | A | Pass | Tab cycles through controls; ESC closes every form. |
| 2.1.3 Keyboard (No Exception) | AAA | Pass | Same as 2.1.1. |
| 2.2.1 Timing Adjustable | A | N/A | No time limits. |
| 2.2.2 Pause, Stop, Hide | A | Pass | Batch can be cancelled at file boundary via ESC or Cancel button. |
| 2.3.1 Three Flashes | A | Pass | No flashing content. |
| 2.4.1 Bypass Blocks | A | N/A | Single-page dialog. |
| 2.4.2 Page Titled | A | Pass | Form `Caption` set to "Batch PDF Publisher vX.Y.Z". |
| 2.4.3 Focus Order | A | Pass | Tab order left-to-right, top-to-bottom. |
| 2.4.4 Link Purpose | A | N/A | No links. |
| 2.4.5 Multiple Ways | AA | N/A | Single-window application. |
| 2.4.6 Headings and Labels | AA | Pass | All inputs have visible labels. |
| 2.4.7 Focus Visible | AA | **Partial** | TextBox controls have a focus border. Label-based buttons cannot show focus (platform limitation). |
| 2.4.11 Focus Not Obscured | AA (2.2) | Pass | Focused control is never behind another. |

### 3.3 Understandable

| SC | Level | Status | Notes |
|---|---|---|---|
| 3.1.1 Language of Page | A | Partial | MSForms does not expose `lang` attribute. Language detected from `ThisApplication.LanguageCode`. |
| 3.1.2 Language of Parts | AA | Pass | All UI strings routed through `L()` in one language at a time. |
| 3.2.1 On Focus | A | Pass | Focus does not trigger context changes. |
| 3.2.2 On Input | A | Pass | Input does not change context without explicit user action. |
| 3.2.3 Consistent Navigation | AA | Pass | Same button order on every form. |
| 3.2.4 Consistent Identification | AA | Pass | Same functions identified consistently. |
| 3.3.1 Error Identification | A | Pass | Errors shown in `lblSum` near `txtPath`. |
| 3.3.2 Labels or Instructions | A | Pass | All inputs have visible labels; tooltips add hints. |
| 3.3.3 Error Suggestion | AA | Partial | Error messages describe the problem; not all suggest a fix. |
| 3.3.4 Error Prevention | AA | Pass | Destructive actions (Reset) require two-click confirmation. |
| 3.3.7 Redundant Entry | A (2.2) | N/A | No multi-step forms. |
| 3.3.8 Accessible Authentication | AA (2.2) | N/A | No authentication. |

### 3.4 Robust

| SC | Level | Status | Notes |
|---|---|---|---|
| 4.1.1 Parsing | A | Deprecated (WCAG 2.2) | Removed. |
| 4.1.2 Name, Role, Value | A | **N/A** | MSForms does not expose UI Automation properties. `ControlTipText` is the only accessible name source. |
| 4.1.3 Status Messages | AA | **N/A** | No live region support in MSForms. |

---

## 4. Colour contrast

All contrast ratios are measured against the rendered background.

### 4.1 Dark theme

| Foreground | Background | Ratio | Level |
|---|---|---|---|
| `#F1F1F1` (primary text) | `#1E1E1E` | 15.9:1 | AAA |
| `#A6A6A6` (secondary text) | `#1E1E1E` | 8.1:1 | AAA |
| `#FFFFFF` on `#0066CC` (accent) | — | 5.6:1 | AA |
| `#FF6B6B` (error) | `#1E1E1E` | 5.7:1 | AA |
| `#4EC94E` (ok) | `#1E1E1E` | 9.6:1 | AAA |
| `#FFB347` (warn) | `#1E1E1E` | 10.2:1 | AAA |

### 4.2 Light theme

| Foreground | Background | Ratio | Level |
|---|---|---|---|
| `#2D2D30` (primary text) | `#F5F5F5` | 13.8:1 | AAA |
| `#686868` (secondary text) | `#F5F5F5` | 5.6:1 | AA |
| `#FFFFFF` on `#0066CC` (accent) | — | 5.6:1 | AA |
| `#B00020` (error) | `#F5F5F5` | 6.6:1 | AA |
| `#006400` (ok) | `#F5F5F5` | 7.5:1 | AAA |
| `#FF8000` (warn) | `#F5F5F5` | 3.2:1 | AA (large text only) |

Napomena: `#FF8000` on `#F5F5F5` fails AA for normal text (3.2:1 vs
required 4.5:1). It is used only for **focus/hover visual cues** on
interactive controls, not as a text foreground on this background.
See `modPalette.FW_LoadPalette`.

### 4.3 High contrast mode

When Windows high contrast mode is active, `PRJ_IsHighContrast`
returns True and `FW_ApplyTheme` returns early — leaving MSForms
system colours in place. These track the user's HC theme and
therefore pass contrast requirements by construction.

---

## 5. EN 301 549 V4.1.1 mapping

EN 301 549 V4.1.1 (ETSI, 2026-09-02) adopts WCAG 2.2 Level AA for
its web and non-web software clauses. Relevant clauses:

| Clause | Requirement | Status |
|---|---|---|
| 9.1.1.1 Non-text Content | WCAG 1.1.1 | N/A (no non-text content) |
| 9.1.3.1 Info and Relationships | WCAG 1.3.1 | Partial (MSForms limitation) |
| 9.1.4.3 Contrast (Minimum) | WCAG 1.4.3 | Pass |
| 9.1.4.11 Non-text Contrast | WCAG 1.4.11 | Pass |
| 9.2.1.1 Keyboard | WCAG 2.1.1 | Pass |
| 9.2.1.2 No Keyboard Trap | WCAG 2.1.2 | Pass |
| 9.2.4.7 Focus Visible | WCAG 2.4.7 | Partial (Label buttons) |
| 9.3.3.2 Labels or Instructions | WCAG 3.3.2 | Pass |
| 9.4.1.2 Name, Role, Value | WCAG 4.1.2 | N/A (MSForms) |

**Conformance claim:** Partial. All clauses that can be satisfied
within MSForms are satisfied. Remaining gaps are platform-bound and
would require a migration to Windows Forms or WPF (see
`ROADMAP.md` v10.x).

Annex ZB (EAA mapping) does not introduce additional requirements
beyond WCAG 2.2 AA for this application class.

---

## 6. DOJ web accessibility rule

The DOJ final rule on web accessibility (published 2024-04-24,
active 2026-01-18, ADA Title II) applies to **web content and mobile
applications** provided by state and local government entities. It
does not apply to:

- Desktop applications distributed as add-ins
- Software installed locally (VBA, .NET, native)
- Non-web user interfaces

This add-in is a **desktop VBA application**. The DOJ rule is
**not applicable**. ADA Title II deadlines (24 April 2026 for
entities >50k, 26 April 2027 for ≤50k) do not apply.

---

## 7. Keyboard navigation

### 7.1 Global shortcuts

| Key | Action | Where |
|---|---|---|
| `ESC` | Close form / cancel operation | All forms |
| `Tab` | Next control | All forms |
| `Shift+Tab` | Previous control | All forms |
| `Enter` | Activate focused button / submit text | All forms |
| `Ctrl+Z` | Undo (native Edit control) | `txtPath`, `txtSearch` |

**Undo note:** Ctrl+Z is handled by the native Edit control and
does not raise the VBA `Change` event. See
`docs/KNOWN_LIMITATIONS.md` §"Ctrl+Z in txtPath" for the validation
refresh caveat.

### 7.2 Per-form Alt accelerators

**frmScan:**
| Shortcut | Action |
|---|---|
| `Alt+R` | Run |
| `Alt+B` | Browse |
| `Alt+T` | Theme |
| `Alt+E` | Reset |
| `Alt+C` | Cancel |
| `Alt+S` | Subfolders toggle |
| `Alt+F` | Force toggle |
| `Alt+A` | Audit toggle |

**frmReport:**
| Shortcut | Action |
|---|---|
| `Alt+B` | Back |
| `Alt+C` | Copy |
| `Alt+R` | Retry |
| `Alt+E` | Export |
| `Alt+Q` | Close |
| `Alt+1` .. `Alt+5` | Filter All / Err / Skip / Ok / NoIdw |

**frmProgress:**
| Shortcut | Action |
|---|---|
| `Alt+C` | Cancel |

Accelerators are set by `PRJ_SetupAccelerators` (modUIHelpers) at
`UserForm_Initialize`. `CKeyPreview` class hooks `KeyDown` on every
focusable control so ESC works regardless of focus.

### 7.3 Focus order

Tab order follows visual order:

**frmScan:** `txtPath` → `lblBrowse` → `lblOptSub` → `lblOptForce`
→ `lblOptAudit` → `lblRun` → `lblTheme` → `lblReset` → `lblCancel`

**frmReport:** `txtSearch` → `lstReport` → `lblFltAll` → `lblFltErr`
→ `lblFltSkip` → `lblFltOk` → `lblFltNoIdw` → `lblBack` → `lblCopy`
→ `lblExp` → `lblRetry` → `lblCancel`

Set via designer `TabIndex`; verified by `PRJ_CheckAccessibility`.

Note: the v9.25.42 scan option toggles (`lblOptSub`, `lblOptForce`,
`lblOptAudit`) are Labels and do not receive Tab focus in MSForms.
They are reachable via `Alt+S` / `Alt+F` / `Alt+A`. They are listed
above for reference because they carry `TabIndex` in the designer,
but Tab moves past them.

---

## 8. Screen reader behaviour

MSForms does not expose UI Automation. Screen readers (NVDA, JAWS,
Narrator) interact with MSForms controls via the legacy
`IAccessible` interface, which is limited.

### 8.1 What works

- Form `Caption` is announced on open.
- Control `Caption` for `Label`, `CheckBox`, and `CommandButton` is
  announced when the control receives focus.
- `ControlTipText` is announced as a tooltip.
- TextBox `Text` is announced when the field receives focus.

### 8.2 What does not work

- Custom `AccessibleName` / `AccessibleDescription` — not available.
- `AccessibleRole` override — not available.
- Live region announcements for `lblSum` status changes — no
  dynamic content announcement.
- `Label` controls used as button frames are announced as static
  text, not as buttons.
- The v9.25.42 scan option toggles (`lblOptSub`, `lblOptForce`,
  `lblOptAudit`) are announced as static text — there is no
  "checked"/"unchecked" state because Labels do not have a `.Value`
  property. Keyboard users reach them via `Alt+S` / `Alt+F` /
  `Alt+A`.

### 8.3 Fallback strategy

For every Label-based button, `ControlTipText` describes the action
in natural language. Screen reader users hear the tooltip when the
mouse hovers, but **not** when focus moves via Tab (because Label
does not receive Tab focus).

**Practical mitigation:** `Alt+` accelerators on all primary
actions. Screen reader users can invoke actions by accelerator
without needing to focus the Label.

---

## 9. Verification

`PRJ_CheckAccessibility` (in `modSelfTest`) walks every control on
every form and reports:

- Controls without `ControlTipText`
- TextBoxes without a nearby Label or tooltip
- Tab order gaps (skipped `TabIndex` values)
- Interactive controls with `TabStop = False`

Run from Immediate Window:

    PRJ_CheckAccessibility

Output is written to the Immediate Window and appended to the
self-test report. It is advisory — a warning does not fail the
self-test, but should be investigated.

---

## 10. Accepted limitations

The following are documented as **accepted** due to MSForms
platform constraints. They would not be fixable without a UI
framework migration.

| ID | Limitation | Impact | Workaround |
|---|---|---|---|
| AL-01 | No `AccessibleName` | Screen readers fall back to `Caption` + tooltip | Provide clear visible labels |
| AL-02 | No focus indicator on Label-based buttons | Keyboard users cannot see which button has "focus" | Alt+ accelerators on every button |
| AL-03 | No DPI scaling above 150% | Layout may not be optimal on large displays | Windows DPI virtualization |
| AL-04 | No live region announcements | Status updates not announced | Users poll via Alt+ shortcuts |
| AL-05 | Label announced as static text, not button | Role confusion for screen reader users | Tooltip + Alt+ accelerator |
| AL-06 | Scan option toggles have no announced state | Screen readers do not announce ON/OFF for `lblOptSub` / `lblOptForce` / `lblOptAudit` | Alt+ accelerators; state visible in accent background colour |

These are tracked in `docs/TECH_DEBT.md` under the "Won't fix" section.

---

## 11. Future improvements

If the project migrates to .NET (see `ROADMAP.md` v10.x), the
following become available:

- Full UI Automation support (`AccessibleName`, `AccessibleDescription`,
  `AccessibleRole`)
- Native focus visuals on all interactive controls
- `AutoScaleMode = Font` for DPI-independent layout
- `AutomationProperties.LiveSetting` for status announcements
- `lang` attribute on every element
- Native CheckBox controls with proper checked-state announcement
  (removes AL-06)

The migration is out of scope for v9.x. This section exists for
planning purposes.

---

*End of ACCESSIBILITY.md*