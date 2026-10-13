# Compliance

Status of relevant standards and regulations for Batch PDF Publisher.

**Version:** 9.25.32
**Last updated:**  2026-10-04
**Host:**  Autodesk Inventor 2027
**Runtime:**  VBA 7.1

## Software license

| Item | Value |
|------|-------|
| Project license | MIT (see LICENSE) |
| Copyright holder | Jugoslav Ilic |
| Copyright year | 2026 |
| Runtime dependencies | Microsoft, Autodesk (proprietary) |
| Redistribution model | Source (.bas/.cls/.frm) or compiled .ivb |

All runtime components are proprietary Microsoft or Autodesk
libraries, used under their respective EULAs. No third-party open
source packages are embedded or linked.

## Accessibility

See `docs/ACCESSIBILITY.md` for the comprehensive accessibility
documentation. This section is a summary.

| Standard | Status | Notes |
|----------|--------|-------|
| WCAG 2.2 AA | Partial (within MSForms) | See ACCESSIBILITY.md §3 |
| ISO/IEC 40500:2025 | Partial | Same as WCAG 2.2 |
| EN 301 549 V4.1.1 | Partial | Adopts WCAG 2.2 AA; see §5 of ACCESSIBILITY.md |
| DOJ web accessibility rule | N/A | Not a web application |
| Section 508 | N/A | Not US federal procurement |

### What passes

- Colour contrast in both themes (Light / Dark): accent #0066CC on
  white = 5.6:1, text #F1F1F1 on #1E1E1E = 15.9:1
- Keyboard navigation: Alt+ accelerators on all buttons
- ESC handler on every form, routed via CKeyPreview class
- Focus ring on txtPath and txtSearch (BorderColor changes on focus)
- Logical tab order (left-to-right, top-to-bottom)
- ControlTipText on all interactive controls (fallback for
  AccessibleName)
- High contrast mode detection — falls back to system colours
- Reset confirmation dialog for destructive actions

### What does not pass (MSForms platform limitations)

- Focus indicator on Label-based buttons — MSForms Label does not
  support custom focus visuals
- Screen reader AccessibleName / AccessibleDescription — MSForms
  does not expose UI Automation properties
- DPI scaling above 150% — MSForms does not have AutoScaleMode
- Live region announcements for dynamic status updates

Full mapping and accepted-limitation list in `docs/ACCESSIBILITY.md`.

### Keyboard shortcuts

| Form | Action | Shortcut |
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

Full details: docs/KNOWN_LIMITATIONS.md and docs/ACCESSIBILITY.md.

## Security

### OWASP Top 10:2025

| Category | Status | Notes |
|----------|--------|-------|
| A01 Broken Access Control | N/A | No web endpoint |
| A02 Security Misconfiguration | Pass | HKCU only, sandbox in %TEMP% |
| A03 Software Supply Chain | Pass | No third-party packages; deps.lock.json + SBOM |
| A04 Cryptographic Failures | N/A | No crypto operations |
| A05 Injection | Pass | Shell calls sanitized |
| A06 Vulnerable Components | Pass | Windows and Autodesk components only |
| A07 Authentication Failures | N/A | No authentication |
| A08 Data Integrity Failures | Pass | Atomic PDF replace with .bak fallback |
| A09 Logging Failures | Pass | Structured log + PII masking + audit trail |
| A10 SSRF | N/A | No server-side requests |

### Data handling

Registry:

- Path: HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs
- Never HKLM. Only HKCU reads and writes.

Sandbox:

- Path: %TEMP%\VBADR\
- Per-user location; requires write access to %TEMP%

Network:

- No outbound network communication
- No telemetry
- No PII collection
- No cloud services

### Log file content

Log entries contain source folder paths. In enterprise deployments
where paths reveal project or client names, treat log files as
confidential. Retention: 24 hours (automatic cleanup at startup),
audit files 90 days.

**PII masking:** LOG_SanitizePII masks %USERPROFILE% and
\Users\<name>\ patterns before writing.

### CVE status

| CVE | Relevant | Notes |
|-----|----------|-------|
| CVE-2026-69439 | No | PDB parsing - not used |
| CVE-2026-71328 | No | PDB parsing - not used |
| CVE-2026-62886 | No | .NET Framework - not used directly |
| CVE-2026-69522 | No | .NET Framework - not used directly |

No known vulnerable components in scope.

## Software Bill of Materials

| Field | Value |
|-------|-------|
| Format | CycloneDX 1.5 (JSON) |
| Location | sbom.cdx.json |
| Manifest | deps.lock.json (human-editable source of truth) |
| Generator | Build-SBOM.ps1 (v2.3.0) |
| Signature | SHA-256 of JSON content |

### Runtime components

| Component | Version | Producer | License | Hash Algorithm |
|-----------|---------|----------|---------|----------------|
| VBA | 7.1 | Microsoft | Proprietary | SHA-256 |
| Inventor API | 2026 / 2027 | Autodesk | Proprietary | SHA-256 |
| stdole | 2.0 | Microsoft | Proprietary | SHA-256 |
| MSForms (FM20) | 2.0 | Microsoft | Proprietary | SHA-256 |

No NuGet, npm, PyPI, or other third-party packages.

### CISA 2026 Minimum Elements

All 23 elements covered — 9 SBOM Metadata, 8 Component Data,
6 Practices & Processes. Element-by-element mapping in
`ARCHITECTURE.md` section 34.4.

### Signature caveat

The `SBOM_Author_Signature_SHA256` field is a self-signature (SHA-256
of JSON content), not an asymmetric author signature. Full CISA 2026
author signature is planned for a future release alongside code
signing.

## Regulatory

### EU AI Act (Regulation EU 2024/1689)

Not applicable. The project uses no AI components. Module modAI was
planned but removed in v9.16.0.

### EU CRA (Regulation EU 2024/2847)

Vulnerability reporting obligation active from 11 September 2026
(24h CSIRT report for actively exploited vulnerabilities). Full
compliance deadline: 11 December 2027.

Compliance artifacts:

- SECURITY.md (vulnerability reporting procedure, CSIRT SLA)
- deps.lock.json (dependency manifest)
- sbom.cdx.json (CycloneDX 1.5)
- CHANGELOG.md (security fixes tracked)
- docs/ROLLBACK.md (recovery procedure)

### GDPR

Not applicable for personal data processing. The project does not
collect or process personal data. Log files may contain file paths;
usernames are masked by LOG_SanitizePII.

### Autodesk API Terms of Service

Applies. The add-in uses the Inventor API and PDF Translator AddIn.
Distribution is subject to the Autodesk EULA.

## Compliance matrix summary

| Area | Status | Blocker |
|------|--------|---------|
| License | Pass | None |
| SBOM | Pass | None |
| Security (OWASP) | Pass | None |
| Accessibility | Partial | MSForms platform limitations |
| AI regulations | N/A | No AI |
| Data protection | Pass | None |
| Supply chain | Pass | deps.lock.json + SBOM |
| EU CRA | Pass | SECURITY.md in place |