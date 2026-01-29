<#
build_mpt.ps1 — Final (Verbose + Logged Failures)

Purpose
-------
Open templates\TemplateBase_WithRibbon.mpt in Microsoft Project (invisible), 
import all VBA modules from macros\production (including .vb converted to .bas), 
and save as templates\ModèleImport.mpt.

Expected Repository Layout (no parameters required)
---------------------------------------------------
<repo-root>\
  scripts\build_mpt.ps1                 # this script (location doesn't matter)
  templates\
    TemplateBase_WithRibbon.mpt         # produced by add_ribbon_to_mpt.ps1
    ModèleImport.mpt                    # output created by this script
  macros\
    production\                         # source of .bas/.cls/.frm and convertible .vb

Key Behavior
------------
- .vb (actually VBA saved as .vb) are converted to .bas (CRLF + ANSI + Attribute VB_Name + Option Explicit).
- Native .bas/.cls/.frm are now normalized to CRLF + ANSI before import; .bas gets VB_Name if missing.
- Non-VBA .vb (likely VB.NET) are silently skipped during conversion (heuristic).
- Per-file import failures emit Write-Warning with the reason (no longer silent).
- Fatal errors (template missing / cannot open / cannot save) fail fast.
- At the end, prints: "Macros imported: X/Y".
- Use -Verbose for candidate-root probing, file enumeration, conversions, and normalization.

Prerequisites
-------------
- Microsoft Project installed.
- Trust Center: Enable "Trust access to the VBA project object model" in Project.

Usage
-----
PS> .\scripts\build_mpt.ps1
PS> .\scripts\build_mpt.ps1 -Verbose

#>

[CmdletBinding()]
param()

# Improve console output for Unicode (paths/messages with accents)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Section {
    param([Parameter(Mandatory)][string]$Text)
    Write-Host ""
    Write-Host "================================="
    Write-Host $Text
    Write-Host "================================="
}

function Release-ComObject {
    param([Parameter(Mandatory)] $ComObj)
    try {
        if ($null -ne $ComObj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObj) }
    } catch { }
}

# Make a valid VBA identifier for module names (letters/digits/_ only; cannot start with digit; avoid spaces/accents)
function New-VbaIdentifier {
    param([Parameter(Mandatory)][string]$BaseName)
    $name = ($BaseName -replace '[^A-Za-z0-9_]', '_')
    if ($name -match '^[0-9]') { $name = "M_$name" }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Module1' }
    if ($name.Length -gt 31) { $name = $name.Substring(0,31) }
    return $name
}

# Convert a .vb (actually VBA) to a valid .bas with header, CRLF, and ANSI encoding
function Convert-VbToBas {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestPath
    )
    $base = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $moduleName = New-VbaIdentifier -BaseName $base

    Write-Verbose ("Converting .vb -> .bas: {0} => {1}" -f (Split-Path -Leaf $SourcePath), (Split-Path -Leaf $DestPath))

    $raw  = Get-Content -LiteralPath $SourcePath -Raw -ErrorAction Stop
    # Normalize endings to CRLF for VBE
    $text = ($raw -replace "`r?`n", "`r`n")

    if ($text -notmatch '(?im)^\s*Attribute\s+VB_Name\s*=') {
        $headerLines = @(
            ('Attribute VB_Name = "{0}"' -f $moduleName)
            'Attribute VB_GlobalNameSpace = False'
            'Attribute VB_Creatable = False'
            'Attribute VB_PredeclaredId = False'
            'Attribute VB_Exposed = False'
            'Option Explicit'
            ''
        )
        $header = ($headerLines -join "`r`n")
        $text = $header + $text
    }

    # VBE prefers ANSI (Default)
    Set-Content -LiteralPath $DestPath -Value $text -Encoding Default
}

# ---------- Resolve project root ----------
Section "Resolving paths"
$ScriptDir = Split-Path -Parent $PSCommandPath
$ParentOfScriptDir = (Get-Item $ScriptDir).Parent.FullName
$candidates = @($ParentOfScriptDir, $ScriptDir, (Get-Location).Path) | Select-Object -Unique
$relativeTemplateIn = 'templates\TemplateBase_WithRibbon.mpt'

$ProjectRoot = $null
foreach ($c in $candidates) {
    $probe = Join-Path $c $relativeTemplateIn
    Write-Verbose "Checking candidate root: $c"
    Write-Verbose " -> Looking for: $probe"
    if (Test-Path -LiteralPath $probe) { $ProjectRoot = $c; break }
}
if (-not $ProjectRoot) {
    $expected = Join-Path $ParentOfScriptDir $relativeTemplateIn
    throw "Missing input template: $expected"
}

$TemplateIn  = Join-Path $ProjectRoot 'templates\TemplateBase_WithRibbon.mpt'
$MacrosRoot  = Join-Path $ProjectRoot 'macros\production'
$TemplateOut = Join-Path $ProjectRoot 'templates\ModèleImport.mpt'
$TempImport  = Join-Path $ScriptDir  '_temp_import_vba'

Write-Host "Resolved ProjectRoot : $ProjectRoot"
Write-Host "TemplateIn           : $TemplateIn"
Write-Host "MacrosRoot           : $MacrosRoot"
Write-Host "TemplateOut          : $TemplateOut"

# ---------- Pre-flight ----------
Section "Pre-flight checks"
if (-not (Test-Path -LiteralPath $TemplateIn)) { throw "Missing input template: $TemplateIn" }
if (-not (Test-Path -LiteralPath $MacrosRoot)) { throw "Missing macros folder: $MacrosRoot" }

# STRICT: pick only supported native VBA files (.bas/.cls/.frm)
$vbaFilesNative = Get-ChildItem -Path $MacrosRoot -Recurse -File |
    Where-Object { $_.Extension -in @('.bas', '.cls', '.frm') }

# STRICT: .vb candidates only
$vbCandidates = Get-ChildItem -Path $MacrosRoot -Recurse -File |
    Where-Object { $_.Extension -ieq '.vb' }

Write-Verbose ("Found native VBA files (.bas/.cls/.frm): {0}" -f (@($vbaFilesNative).Count))
Write-Verbose ("Found .vb candidates: {0}" -f (@($vbCandidates).Count))

# ---------- Normalize native files to ANSI + CRLF (and ensure VB_Name for .bas) ----------
$TempImportNative = Join-Path $ScriptDir '_temp_import_native'
if (Test-Path -LiteralPath $TempImportNative) { Remove-Item -LiteralPath $TempImportNative -Recurse -Force }
$null = New-Item -ItemType Directory -Force -Path $TempImportNative

$normalizedNative = @()

foreach ($src in $vbaFilesNative) {
    try {
        $dest = Join-Path $TempImportNative $src.Name
        Write-Verbose ("Normalizing native VBA file: {0} -> {1}" -f $src.FullName, $dest)

        $raw  = Get-Content -LiteralPath $src.FullName -Raw -ErrorAction Stop

        # Normalize to CRLF
        $text = ($raw -replace "`r?`n", "`r`n")

        if ($src.Extension -ieq '.bas') {
            # Ensure Attribute VB_Name exists
            if ($text -notmatch '(?im)^\s*Attribute\s+VB_Name\s*=') {
                $baseName   = [System.IO.Path]::GetFileNameWithoutExtension($src.Name)
                $moduleName = New-VbaIdentifier -BaseName $baseName
                $headerLines = @(
                    ('Attribute VB_Name = "{0}"' -f $moduleName)
                    'Attribute VB_GlobalNameSpace = False'
                    'Attribute VB_Creatable = False'
                    'Attribute VB_PredeclaredId = False'
                    'Attribute VB_Exposed = False'
                    'Option Explicit'
                    ''
                )
                $header = ($headerLines -join "`r`n")
                $text = $header + $text
            }
        }

        # Save with ANSI for VBE compatibility
        Set-Content -LiteralPath $dest -Value $text -Encoding Default
        $normalizedNative += Get-Item -LiteralPath $dest
    }
    catch {
        Write-Warning ("Failed to normalize native VBA file: {0} - {1}" -f $src.FullName, $_.Exception.Message)
        # Skip this file; it will not be imported
        continue
    }
}

# ---------- Convert .vb -> .bas with proper header/encoding ----------
$convertedBas = @()
if ($vbCandidates -and @($vbCandidates).Count -gt 0) {
    Section "Pre-processing .vb files (copy as .bas)"
    if (Test-Path -LiteralPath $TempImport) { Remove-Item -LiteralPath $TempImport -Recurse -Force }
    $null = New-Item -ItemType Directory -Force -Path $TempImport

    foreach ($vb in $vbCandidates) {
        # Skip likely VB.NET (heuristic)
        $looksDotNet = $false
        try {
            $sample = Get-Content -LiteralPath $vb.FullName -TotalCount 60 -ErrorAction SilentlyContinue
            if ($sample -and ($sample -match '(?i)\b(Imports\s+System|Public\s+Class|End\s+Class|Namespace\s+|End\s+Namespace|Using\s+System)\b')) {
                $looksDotNet = $true
            }
        } catch { }

        if ($looksDotNet) {
            Write-Verbose ("Skipping non-VBA .vb (looks like .NET): {0}" -f $vb.FullName)
            continue
        }

        $destBas = Join-Path $TempImport ("{0}.bas" -f [System.IO.Path]::GetFileNameWithoutExtension($vb.Name))
        try {
            Convert-VbToBas -SourcePath $vb.FullName -DestPath $destBas
            Write-Host ("Converted for import: {0} -> {1}" -f $vb.Name, (Split-Path -Leaf $destBas))
            $convertedBas += Get-Item -LiteralPath $destBas
        }
        catch {
            # Log and skip conversion failure (file-specific; does not stop the build)
            Write-Warning ("Failed to convert .vb to .bas: {0} - {1}" -f $vb.FullName, $_.Exception.Message)
            continue
        }
    }
}

# Final import set = normalized native (.bas/.cls/.frm) + converted .bas (UNIQUE; never any .vb)
$filesToImport = @($normalizedNative) + @($convertedBas)
$filesToImport = $filesToImport | Sort-Object -Property FullName -Unique
$totalExpected = @($filesToImport).Count

if ($totalExpected -eq 0) {
    throw "No VBA files to import. Expected .bas/.cls/.frm or convertible .vb under: $MacrosRoot"
}

# Prepare output directory and overwrite behavior
$null = New-Item -ItemType Directory -Force -Path (Split-Path -Parent $TemplateOut)
if (Test-Path -LiteralPath $TemplateOut) {
    Write-Host "Removing existing output: $TemplateOut"
    Remove-Item -LiteralPath $TemplateOut -Force
}

# ---------- Launch Microsoft Project ----------
Section "Launching Microsoft Project (invisible)"
$projApp  = $null
$vbProj   = $null
$imported = 0
$failedImports = New-Object System.Collections.Generic.List[object]

try {
    $projApp = New-Object -ComObject 'MSProject.Application'
    $projApp.Visible       = $false
    $projApp.DisplayAlerts = $false

    Write-Host "Opening: $TemplateIn"
    [void]$projApp.FileOpen($TemplateIn)

    # Access VBE (requires Trust Center setting)
    try {
        $vbProj = $projApp.VBE.VBProjects.Item(1)
    }
    catch {
        throw "Cannot access the VBA project. Enable 'Trust access to the VBA project object model' in Project Trust Center."
    }

    # ---------- Import macros (LOG WARNINGS on errors) ----------
    Section ("Importing macros (count: {0})" -f $totalExpected)
    foreach ($file in $filesToImport) {
        try {
            [void]$vbProj.VBComponents.Import($file.FullName)
            $imported++
            Write-Host ("Imported: {0}" -f $file.Name)
        }
        catch {
            $msg = $_.Exception.Message
            Write-Warning ("Failed to import: {0} - {1}" -f $file.FullName, $msg)
            $failedImports.Add([pscustomobject]@{ File=$file.FullName; Error=$msg }) | Out-Null
            continue
        }
    }

    # ---------- Save ----------
    Section "Saving output"
    Write-Host "Saving as: $TemplateOut"
    [void]$projApp.FileSaveAs($TemplateOut)

    if (-not (Test-Path -LiteralPath $TemplateOut)) {
        throw "Save failed: Output not found after save: $TemplateOut"
    }

    # ---------- Close ----------
    Section "Closing"
    Write-Host "Closing all documents..."
    [void]$projApp.FileCloseAll()

    Section "Success"
    Write-Host ("Macros imported: {0}/{1}" -f $imported, $totalExpected)

    if ($failedImports.Count -gt 0) {
        Write-Host ""
        Write-Host "Import failures (details already logged as warnings):"
        foreach ($f in $failedImports) {
            Write-Host (" - {0}" -f $f.File)
        }
    }

    Write-Host ("Output file   : {0}" -f $TemplateOut)
}
catch {
    # Only fatal errors (open/save/close/VBE) are surfaced; per-file import errors are warnings
    Write-Error $_
    throw
}
finally {
    if ($projApp) { try { $projApp.Quit() } catch { } }
    Release-ComObject -ComObj $vbProj
    Release-ComObject -ComObj $projApp
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    # Cleanup temp folders
    foreach ($tmp in @($TempImport, $TempImportNative)) {
        if ($tmp -and (Test-Path -LiteralPath $tmp)) {
            try { Remove-Item -LiteralPath $tmp -Recurse -Force } catch { }
        }
    }
}