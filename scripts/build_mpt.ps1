<#
build_mpt.ps1 (Final Silent Version)

- Open templates\TemplateBase_WithRibbon.mpt in Microsoft Project (invisible)
- Import VBA files (.bas/.cls/.frm) from .\macros\production recursively
- Convert only .vb (VBA saved with wrong extension) to valid .bas:
    * Adds Attribute VB_Name header + Option Explicit
    * Normalizes CRLF line endings
    * Writes ANSI (Default code page) for VBE compatibility
- Save as templates\ModèleImport.mpt
- Fail fast for setup/open/save issues
- SILENTLY SKIP any import failures (no prompts, no error messages)
- Log successful imports and final macro count
- Clean close + proper COM release
#>

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
    try { if ($null -ne $ComObj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObj) } } catch { }
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
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestPath
    )
    $base = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $moduleName = New-VbaIdentifier -BaseName $base

    $raw  = Get-Content -LiteralPath $SourcePath -Raw -ErrorAction Stop
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
    Write-Host "Checking candidate root: $c"
    Write-Host " -> Looking for: $probe"
    if (Test-Path -LiteralPath $probe) { $ProjectRoot = $c; break }
}
if (-not $ProjectRoot) {
    $expected = Join-Path $ParentOfScriptDir $relativeTemplateIn
    throw "Missing input template: $expected"
}

$TemplateIn  = Join-Path $ProjectRoot 'templates\TemplateBase_WithRibbon.mpt'
# ✅ Use macros\production as requested
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

# Convert .vb -> .bas with proper header/encoding
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
            # Silent skip of non-VBA .vb
            continue
        }

        $destBas = Join-Path $TempImport ("{0}.bas" -f [System.IO.Path]::GetFileNameWithoutExtension($vb.Name))
        try {
            Convert-VbToBas -SourcePath $vb.FullName -DestPath $destBas
            Write-Host ("Converted for import: {0} -> {1}" -f $vb.Name, (Split-Path -Leaf $destBas))
            $convertedBas += Get-Item -LiteralPath $destBas
        }
        catch {
            # Silent skip on conversion failure (no prompt, no error)
            continue
        }
    }
}

# Final import set = native (.bas/.cls/.frm) + converted .bas (UNIQUE; never any .vb)
# ✅ Force arrays in concatenation and in count check to avoid 'op_Addition' / 'Count' issues
$filesToImport = @($vbaFilesNative) + @($convertedBas)
$filesToImport = $filesToImport | Sort-Object -Property FullName -Unique

if (@($filesToImport).Count -eq 0) {
    throw "No VBA files to import. Expected .bas/.cls/.frm or convertible .vb under: $MacrosRoot"
}

# Prepare output directory and overwrite behavior
$null = New-Item -ItemType Directory -Force -Path (Split-Path -Parent $TemplateOut)
if (Test-Path -LiteralPath $TemplateOut) {
    Write-Host "Removing existing output: $TemplateOut"
    Remove-Item -LiteralPath $TemplateOut -Force
}

# ---------- Launch MS Project ----------
Section "Launching Microsoft Project (invisible)"
$projApp  = $null
$vbProj   = $null
$imported = 0

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

    # ---------- Import macros (SILENT SKIP on errors) ----------
    Section ("Importing macros (count: {0})" -f (@($filesToImport).Count))
    foreach ($file in $filesToImport) {
        try {
            [void]$vbProj.VBComponents.Import($file.FullName)
            $imported++
            Write-Host ("Imported: {0}" -f $file.Name)
        }
        catch {
            # Silent skip: no prompt, no error output
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
    Write-Host ("Macros imported: {0}" -f $imported)
    Write-Host ("Output file   : {0}" -f $TemplateOut)
}
catch {
    # Only fatal errors (open/save/close) are surfaced; per-file import errors are silent
    Write-Error $_
    throw
}
finally {
    if ($projApp) { try { $projApp.Quit() } catch { } }
    Release-ComObject -ComObj $vbProj
    Release-ComObject -ComObj $projApp
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if (Test-Path -LiteralPath $TempImport) {
        try { Remove-Item -LiteralPath $TempImport -Recurse -Force } catch { }
    }
}