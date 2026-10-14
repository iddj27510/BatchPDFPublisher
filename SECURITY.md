# SECURITY.md - VBA-DR-01 Batch PDF Publisher

**Version:** 9.25.42
**Last updated:** 2026-10-04

Vulnerability reporting, EU CRA compliance, OWASP mapping, and SBOM
status.

---

## Supported Versions

| Version | Supported |
|---|---|
| 9.19.x  | Yes |
| 9.18.x  | Yes |
| 9.17.x  | No (superseded) |
| < 9.17  | No |

Security fixes are backported only to the latest minor version.

---

## Reporting a Vulnerability

Report to the maintainer via:
- GitHub Security Advisory (preferred)
- Direct email (engilic@gmail.com)

Include:
- Version affected
- Description of vulnerability
- Steps to reproduce
- Impact assessment
- Suggested fix (if any)

**Do not** open a public GitHub issue for security vulnerabilities.

---

## Severity Response SLA (EU CRA Compliant)

The EU Cyber Resilience Act (Regulation EU 2024/2847) mandates
vulnerability reporting from 11 September 2026.

| Severity | CVSS | Initial response | Patch |
|---|---|---|---|
| Actively exploited | Any | < 24h (CSIRT report) | < 7 days |
| Critical | >= 9.0 | < 48h | < 7 days |
| High | 7.0 - 8.9 | < 5 days | < 30 days |
| Medium | 4.0 - 6.9 | < 10 days | < 90 days |
| Low | < 4.0 | < 30 days | Next release |

**CSIRT reporting:** For actively exploited vulnerabilities, the
maintainer files a report with the relevant national CSIRT within
24 hours of confirmation.

---

## Vulnerability Handling Process

1. **Reception** - Report received and acknowledged within SLA.
2. **Triage** - Confirm, classify by CVSS, assign priority.
3. **Analysis** - Identify root cause, scope, affected versions.
4. **Fix** - Develop and test patch.
5. **Disclosure**:
   - GitHub Security Advisory
   - CHANGELOG.md entry
   - Direct notification to reporter
6. **CSIRT notification** - For actively exploited vulnerabilities.
7. **Post-mortem** - Document in `docs/LESSONS_LEARNED.md`.

---

## Dependency Vulnerability Scanning

Dependencies are declared in `deps.lock.json` and hashed in
`sbom.cdx.json`. Neither file is hand-edited; both are derived from
`Build-SBOM.ps1`. See `ARCHITECTURE.md` section 34 for the full
dependency management workflow.

Since this project uses only Windows-native COM components and
Autodesk APIs, traditional dependency scanning (Dependabot, Snyk)
does not apply. Instead:

- **Windows Update** handles OS-level CVEs (COM runtime).
- **Autodesk patches** handle Inventor API CVEs.
- **Monthly manual check** of Microsoft Security Advisories for:

| Component | Source |
|---|---|
| Scripting.FileSystemObject | Windows Update |
| Scripting.Dictionary | Windows Update |
| System.Collections.ArrayList | Windows Update |
| ADODB.Stream | Windows Update |
| Shell.Application | Windows Update |
| htmlfile | Windows Update |
| OLE DataObject | Windows Update |
| WScript.Shell | Windows Update |
| Inventor API | Autodesk |
| Apprentice Server | Autodesk |
| PDF Translator AddIn | Autodesk |

---

## OWASP Top 10:2025 Mapping

| Category | Applicability | Mitigation |
|---|---|---|
| A01 Broken Access Control | N/A | No authentication, single-user |
| A02 Security Misconfiguration | Applies | Registry under HKCU, no elevation, defaults documented in README |
| A03 Software Supply Chain Failures | Applies | deps.lock.json + SBOM (CycloneDX 1.5); deps limited to Windows-native COM + Autodesk |
| A04 Cryptographic Failures | N/A | No cryptography used |
| A05 Injection | Applies | SN_SanitizePath; no SQL, no XML parsing (see XXE section) |
| A06 Vulnerable Components | Applies | Windows Update + monthly advisory check |
| A07 Authentication Failures | N/A | No authentication |
| A08 Software and Data Integrity | Partially | SBOM hashes; code signing planned for a future release |
| A09 Logging and Monitoring Failures | Partially | Structured logging + PII masking + audit trail; no SIEM |
| A10 Server-Side Request Forgery | N/A | No HTTP requests |

---

## SBOM (Software Bill of Materials)

Generated with `Build-SBOM.ps1` (v2.3.0) in **CycloneDX 1.5** format.

**Location:** `sbom.cdx.json` (repo root)
**Manifest:** `deps.lock.json` (human-editable source of truth)

**CISA 2026 Minimum Elements:**

| Group | Elements | Status |
|---|---|---|
| SBOM Metadata | 9 | Complete |
| Component Data | 8 | Complete, hashes mandatory |
| Practices & Processes | 6 | Complete, structured known unknowns |

**Data Format:** CycloneDX (SWID not accepted per CISA 2026).

**Hash enforcement:** `Build-SBOM.ps1` fails if any component hash
cannot be resolved, unless `-AllowMissingHashes` is passed. Releases
must be generated without that flag.

**Regeneration:** Run `Build-SBOM.ps1` on every release. Attach result
to GitHub Release.

**Signature caveat:** The SBOM currently carries a self-signature
(SHA-256 of JSON content). Full asymmetric author signature is
planned for a future release alongside code signing.

---

## Logging and Audit

Sensitive data is handled as follows:

- **PII masking:** `LOG_SanitizePII` masks `%USERPROFILE%` and
  `\Users\<name>\` patterns before writing to the log. Log lines
  contain `<userprofile>` and `\Users\<user>\` instead of real
  usernames.
- **Non-ASCII handling:** `LOG_Asciify` replaces every character
  above 126 with `?`, except `\t`, `\n`, `\r`, and the middle dot
  `·` (U+00B7) which is used as a visual separator. This ensures log
  files are pure ASCII and readable on any locale.
- **Audit trail:** Events tagged with context `"AUDIT"` are mirrored
  to a separate file `%TEMP%\VBADR\BatchPDF_Audit_<PID>.txt` with
  90-day retention, distinct from the 24h main log. Audit events
  currently cover user preference resets, PDF backup deletion, PDF
  rollback, and orphan cleanup. See `ARCHITECTURE.md` section 35.4.
- **Log location:** `%TEMP%\VBADR\`. Per-user location, no
  cross-user access. Files are not encrypted at rest; the sandbox
  inherits the OS's file permission model.

Do not log raw registry keys, network credentials, or anything else
that would not survive a security review. See `CONTRIBUTING.md`
section "Logging" for the full policy.

---

## HTTPS

All external URLs use HTTPS:

| URL | Purpose |
|---|---|
| `https://www.autodesk.com/support/technical/article/...` | Apprentice Server install guide (modHelp) |

No other network calls in the codebase. No HTTP, no unencrypted
WebSocket, no TCP sockets.

---

## XXE Protection

**Current state:** Project does not parse XML. No XXE risk today.

**Future-proofing:** If XML parsing is added (config file, report
format, Vault communication), follow these rules:

- Use `MSXML2.DOMDocument60` (not `DOMDocument`)
- Set `ProhibitDTD = True` before any `Load` or `LoadXML` call
- Set `ResolveExternals = False`
- Set `ValidateOnParse = False`
- Never parse untrusted XML with default MSXML settings

Example:

    Dim oDoc As Object
    Set oDoc = CreateObject("MSXML2.DOMDocument60")
    oDoc.setProperty "ProhibitDTD", True
    oDoc.setProperty "ResolveExternals", False
    oDoc.async = False
    oDoc.Load sPath

See `docs/ARCHITECTURE.md` section 26 for related N/A items.

---

## Contact

For security reports, open a GitHub Security Advisory or contact the
maintainer. Response times per SLA table above.

---

*End of SECURITY.md*