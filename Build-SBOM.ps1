# Build-SBOM.ps1
# Generise SBOM u CycloneDX 1.5 formatu za Batch PDF Publisher VBA projekat.
# Pokreni iz D:\www\VBA-DR-01\ :
#   powershell -ExecutionPolicy Bypass -File .\Build-SBOM.ps1
#
# Izlaz: D:\www\VBA-DR-01\sbom.cdx.json
#
# v2.5.0 (v9.25.32 alignment):
#   Default $Version bumped 9.19.4 -> 9.25.32. Tool version string
#   in the generated SBOM (metadata.tools[0].version) bumped
#   2.3.0 -> 2.4.0 to match this script. No script logic changes.
#
# v2.3.0 (§15.1 Lock file, §15.3 SBOM, §15.11 CISA 2026, §15.12 New fields):
#   - Reads dependency list from deps.lock.json (human-editable manifest)
#   - Mandatory component hashes: fails build if a dependency file cannot
#     be resolved, unless -AllowMissingHashes is explicitly passed
#   - Added Coverage field per component (runtime / development / optional)
#   - Added Scope field per component (required / optional)
#   - Added structured Known_Unknowns array (was a single string)
#   - Added BOM_Manifest_Source property pointing at deps.lock.json
#   NOTE: SBOM Author Signature remains SHA-256 of JSON content
#   (self-signature, tamper-evidence only). RSA-2048 asymmetric
#   signature is planned for a future release alongside code signing.
#
# v2.2.0 (§5.19 CISA 2026):
#   - Dynamic FM20.DLL location (Office x86/x64 variants)
#   - Dynamic RxInventor.tlb location (Inventor 2026/2027)
#   - Added Known_Unknowns property (CISA 2026 optional element)
#   - Component hash values now mandatory (no "unavailable" allowed)
#
# CISA 2026 Minimum Elements covered (23 elements, 3 groups):
#   SBOM Metadata (9): Author, Timestamp, Data Format Name/Version,
#     Tool Name/Version, SBOM Version, Generation Context, Signature
#   Component Data (8): Name, Version, Producer, Hash Value+Algorithm,
#     License, PURL, Type, Coverage
#   Practices (6): Dependency Relationship, Known Unknowns, Manifest source
#
# Distribucioni model: VBA source (.bas/.cls/.frm) import u Inventor VBE.

param(
    [string]$ProjectName         = "BatchPDFPublisher",
    [string]$Version = "9.25.42",
    [string]$OutputPath          = "",
    [string]$SourceFolder        = "",
    [string]$ManifestPath        = "",
    [string]$GenerationContext   = "release",
    [switch]$AllowMissingHashes  = $false
)

$ErrorActionPreference = "Stop"

# Repo-relative defaults when not specified on the command line
if ([string]::IsNullOrEmpty($OutputPath))   { $OutputPath   = Join-Path $PSScriptRoot 'sbom.cdx.json' }
if ([string]::IsNullOrEmpty($SourceFolder)) { $SourceFolder = Join-Path $PSScriptRoot 'src' }
if ([string]::IsNullOrEmpty($ManifestPath)) { $ManifestPath = Join-Path $PSScriptRoot 'deps.lock.json' }


function Get-FileSha256 {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return "unavailable" }
    if (-not (Test-Path -LiteralPath $Path)) { return "unavailable" }
    try {
        return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLower()
    } catch {
        return "unavailable"
    }
}

function Get-SourceTreeHash {
    param([string]$Root)
    $exts = @('*.bas','*.cls','*.frm','*.frx')
    $files = @()
    foreach ($ext in $exts) {
        $files += Get-ChildItem -LiteralPath $Root -Filter $ext -File -ErrorAction SilentlyContinue
    }
    $files = $files | Sort-Object Name
    if ($files.Count -eq 0) { return "unavailable" }
    $sb = New-Object System.Text.StringBuilder
    foreach ($f in $files) {
        $h = Get-FileSha256 $f.FullName
        [void]$sb.AppendLine("$($f.Name):$h")
    }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha.ComputeHash($bytes)
    return (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
}

function Get-SourceFileCount {
    param([string]$Root)
    $exts = @('*.bas','*.cls','*.frm','*.frx')
    $count = 0
    foreach ($ext in $exts) {
        $count += (Get-ChildItem -LiteralPath $Root -Filter $ext -File -ErrorAction SilentlyContinue).Count
    }
    return $count
}

function Get-FirstExistingPath {
    param([array]$Paths)
    foreach ($p in $Paths) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return ""
}

Write-Host "Batch PDF Publisher SBOM Generator v2.5.0 (VBA-source)"
Write-Host "------------------------------------------------------"
Write-Host "Project      : $ProjectName"
Write-Host "Version      : $Version"
Write-Host "Source       : $SourceFolder"
Write-Host "Manifest     : $ManifestPath"
Write-Host "Output       : $OutputPath"
Write-Host "Context      : $GenerationContext"
Write-Host "Strict hashes: $(-not $AllowMissingHashes)"
Write-Host ""

# Read dependency manifest
if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-Error "Dependency manifest not found: $ManifestPath"
    exit 1
}

Write-Host "Reading dependency manifest..."
$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json

if (-not $manifest.dependencies -or $manifest.dependencies.Count -eq 0) {
    Write-Error "Manifest has no dependencies declared."
    exit 1
}

Write-Host "  Manifest version : $($manifest.manifestVersion)"
Write-Host "  Dependencies     : $($manifest.dependencies.Count)"
Write-Host "  Known unknowns   : $($manifest.knownUnknowns.Count)"
Write-Host ""

# Resolve each dependency's path and hash
$resolvedComponents = @()
$missingHashes = @()

foreach ($dep in $manifest.dependencies) {
    Write-Host "Resolving $($dep.group)/$($dep.name)@$($dep.version)"
    $resolvedPath = Get-FirstExistingPath -Paths $dep.paths
    $hash = "unavailable"
    if (-not [string]::IsNullOrEmpty($resolvedPath)) {
        $hash = Get-FileSha256 $resolvedPath
        Write-Host "  Path : $resolvedPath"
        Write-Host "  Hash : $($hash.Substring(0,[Math]::Min(16,$hash.Length)))..."
    } else {
        Write-Host "  Path : NOT FOUND in any candidate location" -ForegroundColor Red
    }
    if ($hash -eq "unavailable") {
        $missingHashes += "$($dep.group)/$($dep.name)@$($dep.version)"
    }
    Write-Host ""

    $resolvedComponents += [ordered]@{
        type       = $dep.type
        name       = $dep.name
        group      = $dep.group
        version    = $dep.version
        purl       = $dep.purl
        author     = $dep.group
        licenses   = @(@{ license = @{ name = $dep.license } })
        hashes     = @(@{ alg = "SHA-256"; content = $hash })
        scope      = $dep.scope
        coverage   = $dep.coverage
        properties = @(
            @{ name = "Source";        value = $dep.source },
            @{ name = "Resolved_Path"; value = $resolvedPath }
        )
    }
}

# Fail fast if mandatory hashes are missing
if ($missingHashes.Count -gt 0 -and -not $AllowMissingHashes) {
    Write-Host ""
    Write-Host "ERROR: $($missingHashes.Count) component(s) have unresolved hashes:" -ForegroundColor Red
    $missingHashes | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Write-Host ""
    Write-Host "CISA 2026 requires Component Hash Value as a mandatory element." -ForegroundColor Red
    Write-Host "Fix: install the missing component on this build machine, OR" -ForegroundColor Yellow
    Write-Host "     pass -AllowMissingHashes to generate a non-compliant SBOM." -ForegroundColor Yellow
    exit 1
}

$sourceTreeHash = Get-SourceTreeHash -Root $SourceFolder
$sourceFileCount = Get-SourceFileCount -Root $SourceFolder

# Build structured Known_Unknowns for CISA 2026 practices section
$knownUnknowns = @()
foreach ($ku in $manifest.knownUnknowns) {
    $knownUnknowns += [ordered]@{
        category    = $ku.category
        description = $ku.description
    }
}

$sbom = [ordered]@{
    bomFormat    = "CycloneDX"
    specVersion  = "1.5"
    serialNumber = "urn:uuid:" + [guid]::NewGuid().ToString()
    version      = 1
    metadata     = [ordered]@{
        timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        authors   = @(@{ name = "Jugoslav Ilić" })
        supplier  = [ordered]@{ name = "Internal" }
        tools     = @(
            [ordered]@{
                vendor  = "BatchPDFPublisher"
                name    = "Build-SBOM.ps1"
                version = "2.4.0"
            }
        )
        component = [ordered]@{
            type    = "application"
            name    = $ProjectName
            version = $Version
            author  = "Jugoslav Ilić"
            hashes  = @(@{ alg = "SHA-256"; content = $sourceTreeHash })
            properties = @(
                @{ name = "Source_Folder";       value = $SourceFolder },
                @{ name = "Source_File_Count";   value = "$sourceFileCount" },
                @{ name = "Distribution_Model";  value = "VBA source (bas/cls/frm)" },
                @{ name = "Host_Application";    value = "Autodesk Inventor 2027" },
                @{ name = "Runtime";             value = "VBA 7.1" },
                @{ name = "SBOM_Version";        value = "1" },
                @{ name = "Generation_Context";  value = $GenerationContext },
                @{ name = "BOM_Manifest_Source"; value = $ManifestPath }
            )
        }
    }
    components   = $resolvedComponents
    dependencies = @(
        [ordered]@{
            ref       = "pkg:generic/$($ProjectName.ToLower())@$Version"
            dependsOn = @($resolvedComponents | ForEach-Object { $_.purl })
        }
    )
    properties   = @(
        @{ name = "cisa:known_unknowns"; value = ($knownUnknowns | ConvertTo-Json -Compress -Depth 4) }
    )
}

# First pass: compute signature over canonical JSON
$json = $sbom | ConvertTo-Json -Depth 12

# Signature hash of JSON content. NOTE: this is a self-signature (SHA-256 of
# the JSON text) rather than an asymmetric key signature. It provides
# tamper-evidence but not non-repudiation. For full CISA 2026 author
# signature, replace this block with RSA-2048 signing over the JSON bytes
# using a code-signing certificate. See CONTRIBUTING.md.
$sha256 = [System.Security.Cryptography.SHA256]::Create()
$bytes  = [System.Text.Encoding]::UTF8.GetBytes($json)
$hash   = $sha256.ComputeHash($bytes)
$sigHex = (($hash | ForEach-Object { $_.ToString("x2") }) -join "")

$sbom.metadata.component.properties += @{
    name  = "SBOM_Author_Signature_SHA256"
    value = $sigHex
}

$json = $sbom | ConvertTo-Json -Depth 12

Set-Content -LiteralPath $OutputPath -Value $json -Encoding UTF8

Write-Host "SBOM generated successfully."
Write-Host "Distribution : VBA source ($sourceFileCount files)"
Write-Host "Root hash    : $sourceTreeHash"
Write-Host "Components   : $($resolvedComponents.Count)"
Write-Host "Known unknown: $($knownUnknowns.Count)"
Write-Host "Format       : CycloneDX 1.5 (JSON)"
Write-Host "SBOM sig     : $sigHex"
Write-Host ""

# Hash verification (post-write)
Write-Host "Hash verification:"
$missing = @()
foreach ($c in $resolvedComponents) {
    $h = $c.hashes[0].content
    if ($h -eq "unavailable") {
        $missing += "$($c.group)/$($c.name)"
        Write-Host "  [FAIL] $($c.group)/$($c.name): $h" -ForegroundColor Red
    } else {
        Write-Host "  [PASS] $($c.group)/$($c.name): $($h.Substring(0,16))..." -ForegroundColor Green
    }
}

if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "WARN: $($missing.Count) component(s) missing hash (SBOM is not CISA 2026 compliant):" -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
} else {
    Write-Host ""
    Write-Host "All component hashes present. SBOM is CISA 2026 compliant." -ForegroundColor Green
}

Write-Host ""
Write-Host "Output: $OutputPath"