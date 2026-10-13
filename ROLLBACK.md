# Rollback Procedure

**Version:** 9.25.32

Procedure for reverting Batch PDF Publisher to a previous version.

## When to rollback

- New version introduces a regression in production
- First-run crash or startup failure
- Corrupted registry settings after upgrade
- Batch processing produces incorrect output

## Pre-rollback checklist

1. Stop all running batches. Close Inventor completely.
2. Export current state for diagnostics:
   - `%TEMP%\VBADR\BatchPDF_Log_*.txt` — copy all log files
   - `%TEMP%\VBADR\BatchPDF_Audit_*.txt` — copy audit files
   - `HKCU\Software\VB and VBA Program Settings\VBADR01\` — export registry
   - `%TEMP%\VBADR\crash_*.txt` — copy if any exist
3. Note current version from Immediate Window:

       ? PRJ_VERSION

## Rollback steps

### Step 1 — Close Inventor

Fully exit Inventor (not just the VBA editor). Verify with Task Manager
that `Inventor.exe` is not running.

### Step 2 — Restore from Git (preferred)

The project is under Git version control. Rollback is a matter of
checking out a previous tag.

List available versions:

    cd D:\www\VBA-DR-01
    git tag --list "v9.*"
    git log --oneline --decorate -20

Switch to a specific tag:

    git checkout v9.19.4

Then re-import the source files into the VBE:

1. Open Inventor
2. Press Alt+F11
3. In Project Explorer, remove all BatchPDF* modules
   (right-click each, Remove, choose No when prompted to export)
4. File -> Import File
5. Select all .bas, .cls, .frm from src\
6. Debug -> Compile VBAProject
7. Run RunAllTests from Immediate Window

Expected: 40 / 40 passed

### Step 3 — Restore from .ivb backup (fallback)

If Git is unavailable, use a timestamped .ivb backup.

Backups live in: D:\www\VBA-DR-01\backups\

To create a new backup:

    cd D:\www\VBA-DR-01
    .\scripts\backup-ivb.ps1

To restore a specific backup:

    Copy-Item "D:\www\VBA-DR-01\backups\Default_<timestamp>.ivb" `
              "C:\Users\Public\Documents\Autodesk\Inventor 2027\Macros\Default.ivb" -Force

Restart Inventor. Run RunAllTests to verify.

### Step 4 — Registry schema

Between v9.14.x and v9.18.5 the registry schema evolved from 1 to 2:

- **v1** — original schema (Theme, Path, Sub, Force, Audit, LogLevel,
  SchemaVersion, Lang)
- **v2** — added `LogJson` key (v9.18.5)

`STATE_MigrateSchema` handles both directions gracefully:

- Upgrading (v1 → v2): no-op, missing `LogJson` defaults to `False`.
- Downgrading (v2 → v1): `LogJson` is ignored by v1 code.
- Unknown future version: logs WARN, leaves registry untouched.

If a rollback fails to restore preferences, run
`ResetAllSettings` from Immediate Window, or manually delete:

    HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs

Restart Inventor and re-enter preferences.

Warning: this deletes all user preferences.

### Step 5 — Sandbox cleanup (optional)

Old temp files may interfere:

1. Delete %TEMP%\VBADR\
2. Sandbox is recreated on next startup

Note: audit files with 90-day retention live in the same sandbox.
Deleting the sandbox also deletes the audit trail. Copy them first if
they are needed for compliance.

## Post-rollback verification

1. Run_BatchPDFPublisher — frmScan opens without error
2. Batch with 3 test IDW files (T01_HappyPath) — all 3 processed
3. Report form opens with correct summary
4. Check log: %TEMP%\VBADR\BatchPDF_Log_<PID>.txt for unexpected errors
5. Check audit: %TEMP%\VBADR\BatchPDF_Audit_<PID>.txt exists if any
   AUDIT events fired
6. Run RunAllTests — expected: 40 / 40 passed

## Backup discipline

### Git tags

Every release gets a Git tag vX.Y.Z. To list:

    git tag --list

### .ivb backups

Run scripts\backup-ivb.ps1 before any major refactoring session.

The script:

- Creates a timestamped copy in backups\
- Prunes backups older than 30 days (configurable)
- Prints a summary

Retention: 30 days default.

### Release artifacts

Each release is archived in releases\ as:

- BatchPDFPublisher-vX.Y.Z.ivb — compiled project
- sbom-vX.Y.Z.cdx.json — SBOM snapshot
- CHANGELOG-vX.Y.Z.md — frozen changelog

## Emergency: crash-on-startup

If rollback is impossible (no Git, no backup):

1. Start Inventor without loading the add-in
   (disable VBA auto-load temporarily)
2. In VBE, remove all BatchPDF* modules
3. If source is lost:
   - Delete registry: HKCU\Software\VB and VBA Program Settings\VBADR01
   - Delete sandbox: %TEMP%\VBADR\
   - The add-in is gone but the system is clean

## Quick reference: verify version

From Immediate Window:

    ? PRJ_VERSION

Expected for this document: 9.25.32

## Cross-references

- docs\ARCHITECTURE.md — module map
- docs\KNOWN_LIMITATIONS.md — limitations and workarounds
- CHANGELOG.md — full version history
- CONTRIBUTING.md — Git workflow and release process
- SECURITY.md — vulnerability reporting