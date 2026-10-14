Attribute VB_Name = "modConstants"
Option Explicit
' Module    : modConstants
' Purpose   : All Public Const declarations for the project.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.9 (Scan options: label toggles):
'   PRJ_VERSION bumped 9.19.8 -> 9.19.9. Aligns with the
'   frmScan scan-option refactor (checkbox -> single toggle
'   labels). See frmScan header for details.
'
' v9.19.8 (Checkbox first-click fix):
'   PRJ_VERSION bumped 9.19.7 -> 9.19.8. Superseded by the
'   v9.19.9 label refactor � the checkbox that needed the
'   workaround is gone.
'
' v9.19.7 (BUG-02 Host version guard fix):
'   Added HOST_IsSupportedMajor(nMajor). Prior revisions
'   compared SoftwareVersion.Major against 26 and 27 believing
'   those were Inventor 2026 and 2027. Inventor uses an offset
'   of 1996: 2026 -> Major 30, 2027 -> Major 31.
'   PRJ_VERSION bumped 9.19.6 -> 9.19.7.
'
' v9.19.6 (Version alignment):
'   PRJ_VERSION bumped 9.19.5 -> 9.19.6.
'
' This module declares constants and one pure mapping function.
' HOST_IsSupportedMajor is stateless and side-effect free � it
' is a lookup, not logic in the traditional sense. Keeping it
' here gives both modEntry and modSelfTest a single source of
' truth for the host version check without introducing a new
' module for four lines of code.
' Every mutable value lives in modState.
'
' Requires  : (none)
' Notes     : No line-continuation characters in this module.
'             No decorative separator comment lines.

' Version and application identity
Public Const PRJ_VERSION  As String = "9.25.42"
Public Const PDF_GUID    As String = "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"

' Path and size limits
Public Const MAX_PATH_LEN As Long = 260
Public Const RELPATH_MAX_LEN As Long = 70
Public Const PATH_MAX_PARENT_ITER As Long = 20

' Timing
Public Const SECONDS_PER_DAY As Long = 86400
Public Const MS_PER_SECOND   As Long = 1000
Public Const SLEEP_CLIP_MS As Long = 300
Public Const UI_SLEEP_CHUNK_MS As Long = 50
Public Const VALIDATE_DOEVENTS_EVERY As Long = 1
Public Const KEY_DEBOUNCE_SEC      As Double = 0.5
Public Const RESET_ARM_TIMEOUT_SEC As Double = 3#
Public Const SCAN_PROGRESS_MIN_INTERVAL_SEC As Double = 0.2

' UI message durations
Public Const UI_MSG_BRIEF_MS   As Long = 900
Public Const UI_MSG_SHORT_MS   As Long = 1200
Public Const UI_MSG_MEDIUM_MS  As Long = 1500
Public Const UI_MSG_LONG_MS    As Long = 1800
Public Const UI_MSG_ERROR_MS   As Long = 2500

' UI font sizing
Public Const UI_FONT_BASE_SIZE As Single = 9
Public Const UI_FONT_MIN_SIZE  As Single = 7
Public Const UI_FONT_PADDING   As Single = 4
Public Const UI_FONT_SHRINK_MAX_ITER As Long = 20

' FIT framework
Public Const FIT_MIN_FONT As Double = 7#
Public Const FIT_SHRINK_STEP As Double = 0.5
Public Const FIT_START_FONT As Double = 9#
Public Const FIT_ELLIPSIS_TAIL_CHARS As Long = 8

' Cache
Public Const CACHE_MAX_ENTRIES As Long = 10000
Public Const CACHE_EVICT_BATCH As Long = 2000

' Logging
Public Const LOG_MAX_BYTES As Long = 1048576
Public Const LOG_FLUSH_EVERY As Long = 50
Public Const LOG_ROTATE_MAX As Long = 5
Public Const LOG_DEBUG    As Long = 0
Public Const LOG_INFO     As Long = 1
Public Const LOG_WARN     As Long = 2
Public Const LOG_ERROR    As Long = 3
Public Const LOG_CRITICAL As Long = 4
Public Const LOG_FILE_PREFIX  As String = "BatchPDF_Log_"
Public Const LOG_AUDIT_PREFIX As String = "BatchPDF_Audit_"

' Orphan cleanup
Public Const ORPHAN_MAX_AGE_HOURS As Long = 24
Public Const ORPHAN_BAK_MAX_AGE_DAYS As Long = 7
Public Const ORPHAN_AUDIT_MAX_AGE_DAYS As Long = 90

' Disk space
Public Const MIN_FREE_DISK_BYTES As Double = 104857600#

' Temp file prefix
Public Const TEMP_FILE_PREFIX As String = "_tmp_"

' PDF export tuning
Public Const PDF_TEMP_SEG_MAX As Long = 40
Public Const PDF_TEMP_RAND_MAX As Long = 10000
Public Const PDF_VECTOR_RESOLUTION As Long = 400
Public Const PDF_ALL_COLOR_AS_BLACK As Boolean = False
Public Const PDF_PDF_STANDARD As String = ""

' Retry policy
Public Const RETRY_MAX_ATTEMPTS As Long = 5
Public Const RETRY_BASE_SEC     As Double = 0.5
Public Const RETRY_BACKOFF_MULT As Double = 2#
Public Const RETRY_MIN_MS       As Long = 100

' QA thresholds
Public Const QA_STAT_IDW_EMPTY_VIEW As String = "[EmptyView]"
Public Const QA_STAT_IDW_NO_TITLE   As String = "[NoTitleBlock]"
Public Const QA_STAT_PDF_TOO_SMALL  As String = "[PDFSmall]"
Public Const QA_PDF_MIN_BYTES As Double = 10240#

' Tag conventions
Public Const TAG_FILTER_PREFIX As String = "btn-filter-"
Public Const TAG_BTN_PREFIX    As String = "btn-"
Public Const TAG_TXT_SUFFIX    As String = "-txt"

' Registry
Public Const REG_SECTION    As String = "VBADR01"
Public Const REG_PREFS      As String = "Prefs"
Public Const REG_KEY_THEME  As String = "Theme"
Public Const REG_KEY_PATH   As String = "Path"
Public Const REG_KEY_SUB    As String = "Sub"
Public Const REG_KEY_FORCE  As String = "Force"
Public Const REG_KEY_AUDIT  As String = "Audit"
Public Const REG_KEY_LOGLEVEL As String = "LogLevel"
Public Const REG_KEY_LOGJSON  As String = "LogJson"
Public Const REG_KEY_SCHEMAVERSION As String = "SchemaVersion"
Public Const REG_KEY_LANG   As String = "Lang"

' Current registry schema version. Bump when the shape of the
' registry tree changes (key renamed, default changed, key removed).
' STATE_MigrateSchema applies migrations from older versions.
Public Const STATE_SCHEMA_VERSION As Long = 2

' Report markers
Public Const RPT_MARK_OK    As String = "[+]"
Public Const RPT_MARK_WARN  As String = "[w]"
Public Const RPT_MARK_ERROR As String = "[!]"
Public Const RPT_MARK_SKIP  As String = "[o]"
Public Const RPT_MARK_INFO  As String = "[i]"

' Status prefixes � PDF export pipeline (modPDFExport emits these)
Public Const STAT_ERROR          As String = "Error"
Public Const STAT_SKIPPED_EMPTY  As String = "Skipped (Empty)"
Public Const STAT_SKIPPED        As String = "Skipped"
Public Const STAT_OVERWRITTEN    As String = "Overwritten"
Public Const STAT_CREATED        As String = "Created"

' Status prefixes � Validate pipeline (PRJ_CheckPDFStatus emits these).
Public Const STAT_MISSING  As String = "Missing"
Public Const STAT_STALE    As String = "Stale"
Public Const STAT_MISMATCH As String = "Mismatch"
Public Const STAT_OK       As String = "OK"

' Error classification
Public Const EC_UNKNOWN       As Long = 0
Public Const EC_NETWORK       As Long = 1
Public Const EC_PERMISSION    As Long = 2
Public Const EC_DATA          As Long = 3
Public Const EC_FILE_LOCK     As Long = 4
Public Const EC_COM           As Long = 5
Public Const EC_USER_CANCEL   As Long = 6

' Host version guard. Inventor uses an offset of 1996 for the
' SoftwareVersion.Major property:
'   2022 -> 26, 2023 -> 27, 2024 -> 28, 2025 -> 29,
'   2026 -> 30, 2027 -> 31
' Single source of truth for the version guard, consumed by
' PRJ_CheckHostVersion (modEntry) and PRJ_SelfTest (modSelfTest).
Public Function HOST_IsSupportedMajor(ByVal nMajor As Long) As Boolean
    HOST_IsSupportedMajor = (nMajor = 30 Or nMajor = 31)
End Function

' END: modConstants v9.19.9
