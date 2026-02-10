<#
build_mpt.ps1 — Final (Verbose + Logged Failures)

Purpose
-------
Open templates\TemplateBase_WithRibbon.mpt in Microsoft Project (invisible),
import all VBA modules from macros\production (including .vb converted to .bas),
and save as templates\ModeleImport.mpt.

Also:
- Extract customUI/customUI14.xml (from base if ZIP, else from templates\customUI14.xml).
- Apply RibbonX to the active project before saving (binary-safe via SetCustomUI).
- Post-build check: re-open output and re-apply SetCustomUI (fail build if missing).

Expected Repository Layout (no parameters required)
---------------------------------------------------
<repo-root>\
  scripts\build_mpt.ps1
  templates\
    TemplateBase_WithRibbon.mpt
    ModeleImport.mpt            # output created by this script
    customUI14.xml              # fallback ribbon XML (used if base is binary)
  macros\
    production\                 # .bas/.cls/.frm and convertible .vb

Key Behavior
------------
- .vb (actually VBA saved as .vb) are converted to .bas (CRLF + ANSI + Attribute VB_Name + Option Explicit).
- Native .bas/.cls/.frm normalized to CRLF + ANSI; .bas gets VB_Name if missing.
- Non-VBA .vb (likely VB.NET) are skipped (heuristic).
- Per-file import failures are warnings; fatal open/save/COM errors fail fast.
- RibbonX (customUI14) is injected via SetCustomUI (pre-save) and verified post-build.
- End of run prints: "Macros imported: X/Y".
- Use -Verbose for details.

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
param(
    [int]$RibbonTimeoutSeconds = 240,
    [int]$RibbonRetries = 1,
    [switch]$AllowRibbonTimeout,
    [switch]$RequireRibbon,
    [switch]$SkipRibbonApply,
    [switch]$UsePostSaveApply
)

# Improve console output for Unicode (paths/messages with accents)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try { $null = Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue } catch { }

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

# --- PATCH D1: Ribbon helpers (extract & tolerant open) ----------------------

function Get-CustomUIXml {
    param(
        [Parameter(Mandatory)][string]$TemplateBaseWithRibbon,
        [Parameter(Mandatory)][string]$FallbackXmlPath
    )
    # Returns the text of customUI14.xml (2009/07) or $null if not found.
    # Strategy:
    # 1) If base is ZIP/Open XML → extract /customUI/customUI14.xml
    # 2) Else if FallbackXmlPath exists → read it
    # 3) Else → $null
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $fs = [System.IO.File]::OpenRead($TemplateBaseWithRibbon)
        $b1 = $fs.ReadByte(); $b2 = $fs.ReadByte(); $fs.Close()
        $isZip = ($b1 -eq 0x50 -and $b2 -eq 0x4B)  # 'PK'
    } catch { $isZip = $false }

    if ($isZip) {
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($TemplateBaseWithRibbon)
            $entry = $zip.Entries | Where-Object { $_.FullName -eq 'customUI/customUI14.xml' }
            if ($entry) {
                $tmp = Join-Path $env:TEMP ("customUI14_{0}.xml" -f ([guid]::NewGuid()))
                $entry.ExtractToFile($tmp, $true)
                $xml = Get-Content -LiteralPath $tmp -Raw -ErrorAction Stop
                Remove-Item -LiteralPath $tmp -Force
                $zip.Dispose()
                return $xml
            }
            $zip.Dispose()
        } catch { }
    }

    if (Test-Path -LiteralPath $FallbackXmlPath) {
        try {
            return (Get-Content -LiteralPath $FallbackXmlPath -Raw -ErrorAction Stop)
        } catch { }
    }
    return $null
}

function Open-ProjectTolerant {
    param(
        [Parameter(Mandatory)][object]$ProjApp,
        [Parameter(Mandatory)][string]$Path
    )
    # Try minimal FileOpenEx signatures first (vary by version), then fall back to FileOpen.
    try { $ProjApp.FileOpenEx($Path, $true); return } catch { }
    try { $ProjApp.FileOpenEx($Path, $true, $false); return } catch { }
    try { $ProjApp.FileOpenEx($Path); return } catch { }
    $ProjApp.FileOpen($Path)
}

function Wait-ActiveProject {
    param([Parameter(Mandatory)][object]$ProjApp,[int]$Seconds=20)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while (-not $ProjApp.ActiveProject -and $sw.Elapsed.TotalSeconds -lt $Seconds) { Start-Sleep -Milliseconds 200 }
    return $ProjApp.ActiveProject
}

# Apply RibbonX in a separate STA PowerShell process with a timeout.
function Apply-RibbonXToFileWithTimeout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$RibbonXml,
        [int]$TimeoutSeconds = 90
    )

    $xmlPath = Join-Path $env:TEMP ("customUI14_{0}.xml" -f ([guid]::NewGuid()))
    $scriptPath = Join-Path $env:TEMP ("apply_ribbon_{0}.ps1" -f ([guid]::NewGuid()))

    try {
        Set-Content -LiteralPath $xmlPath -Value $RibbonXml -Encoding UTF8

$script = @'
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$XmlPath
)

$null = Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

$app = $null
try {
    $app = New-Object -ComObject 'MSProject.Application'
    try {
        # 3 = msoAutomationSecurityForceDisable
        $app.AutomationSecurity = 3
    } catch { }
    # Make UI visible to avoid SetCustomUI hang in headless mode
    $app.Visible       = $true
    $app.DisplayAlerts = $false
    $app.FileOpen($Path)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while (-not $app.ActiveProject -and $sw.Elapsed.TotalSeconds -lt 20) {
        Start-Sleep -Milliseconds 200
        try { [System.Windows.Forms.Application]::DoEvents() } catch { }
    }
    if (-not $app.ActiveProject) { throw "Timed out waiting for ActiveProject." }
    $xml = Get-Content -LiteralPath $XmlPath -Raw
    $app.ActiveProject.SetCustomUI($xml)
    [void]$app.FileSave()
    [void]$app.FileCloseAll()
} finally {
    if ($app) { try { $app.Quit() } catch { } }
}
'@

        Set-Content -LiteralPath $scriptPath -Value $script -Encoding UTF8

        $proc = Start-Process -FilePath "powershell" -ArgumentList @(
            "-NoProfile",
            "-STA",
            "-File", $scriptPath,
            "-Path", $FilePath,
            "-XmlPath", $xmlPath
        ) -PassThru
        $waitMs = [Math]::Max(1000, $TimeoutSeconds * 1000)
        if (-not $proc.WaitForExit($waitMs)) {
            try { Stop-Process -Id $proc.Id -Force } catch { }
            throw "SetCustomUI timed out after $TimeoutSeconds seconds."
        }
    }
    finally {
        if (Test-Path -LiteralPath $xmlPath) { Remove-Item -LiteralPath $xmlPath -Force }
        if (Test-Path -LiteralPath $scriptPath) { Remove-Item -LiteralPath $scriptPath -Force }
    }
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

# --- PATCH A: Macros root resolution (prefer production, fallback to macros if missing or empty)
$MacrosRootPrimary   = Join-Path $ProjectRoot 'macros\production'
$MacrosRootFallback  = Join-Path $ProjectRoot 'macros'
if (-not (Test-Path -LiteralPath $MacrosRootPrimary)) {
    $MacrosRoot = $MacrosRootFallback
} else {
    $candidateCount = @(Get-ChildItem -Path $MacrosRootPrimary -File -ErrorAction SilentlyContinue).Count
    if ($candidateCount -eq 0) {
        $MacrosRoot = $MacrosRootFallback
    } else {
        $MacrosRoot = $MacrosRootPrimary
    }
}

# --- Output name standardized to ASCII per request
$TemplateOut = Join-Path $ProjectRoot 'templates\ModeleImport.mpt'
$TempImport  = Join-Path $ScriptDir  '_temp_import_vba'

Write-Host "Resolved ProjectRoot : $ProjectRoot"
Write-Host "TemplateIn           : $TemplateIn"
Write-Host "MacrosRoot           : $MacrosRoot"
Write-Host "TemplateOut          : $TemplateOut"

# Fallback Ribbon XML path if base is binary
$CustomUiXmlFallback = Join-Path $ProjectRoot 'templates\customUI14.xml'

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
    try {
        # 3 = msoAutomationSecurityForceDisable
        $projApp.AutomationSecurity = 3
    } catch { }
    $projApp.Visible       = $false
    $projApp.DisplayAlerts = $false

    Write-Host "Opening: $TemplateIn"
    [void]$projApp.FileOpen($TemplateIn)

    # Access VBE (requires Trust Center setting)
    try {
        $vbProj = $projApp.VBE.ActiveVBProject
    }
    catch {
        throw "Cannot access the VBA project. Enable 'Trust access to the VBA project object model' in Project Trust Center."
    }

    # --- PATCH B: Purge existing modules (keep ThisProject only) for deterministic builds
    Section "Purging existing VBA modules"
    for ($i = $vbProj.VBComponents.Count; $i -ge 1; $i--) {
        $comp  = $vbProj.VBComponents.Item($i)
        $name  = $comp.Name  # capture before removal
        if ($name -ne 'ThisProject') {
            try {
                $vbProj.VBComponents.Remove($comp)
                Write-Host ("Removed module: {0}" -f $name)
            } catch {
                Write-Warning ("Failed to remove module {0}: {1}" -f $name, $_.Exception.Message)
            }
        }
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

    # 🔴 CRITICAL: Fail the build if nothing imported
    if ($imported -eq 0 -and $totalExpected -gt 0) {
        Write-Host "[FAIL] CRITICAL: No macros imported ($totalExpected expected)" -ForegroundColor Red
        Write-Host "Check macro files in /macros/production/" -ForegroundColor Yellow
        exit 1
    }

    # --- PATCH C (fixed): Verify required RibbonX callbacks before saving
    Section "Validating RibbonX callbacks"
    $found = @{
        OnRibbonLoad      = $false
        GenerateDashboard = $false
    }
    foreach ($comp in $projApp.VBE.ActiveVBProject.VBComponents) {
        # 1 = Standard module; 2 = Class module; 3 = Form; 100 = Document/ThisProject
        if ($comp.Type -eq 1) {
            $code = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
            if ($code -match 'Public\s+Sub\s+OnRibbonLoad\s*\(')      { $found['OnRibbonLoad']      = $true }
            if ($code -match 'Public\s+Sub\s+GenerateDashboard\s*\(') { $found['GenerateDashboard'] = $true }
        }
    }
    $missing = @()
    foreach ($k in $found.Keys) { if (-not $found[$k]) { $missing += $k } }
    if ($missing.Count -gt 0) {
        throw "Missing required callbacks: $($missing -join ', ')"
    } else {
        Write-Host "All required RibbonX callbacks found."
    }

    # --- PATCH C2: Ensure ThisProject has public callback wrappers (Project sometimes resolves here first)
    Section "Ensuring ThisProject callback wrappers"
    try {
        $tp = $vbProj.VBComponents.Item("ThisProject")
    } catch {
        $tp = $null
    }
    if ($tp) {
        try {
            $tpCodeModule = $tp.CodeModule
            $tpCode = $tpCodeModule.Lines(1, $tpCodeModule.CountOfLines)

            if ($tpCode -notmatch 'Public\s+Sub\s+OnRibbonLoad\s*\(') {
                $code = "Public Sub OnRibbonLoad(ByVal ribbon As Object)`r`n" +
                        "    On Error Resume Next`r`n" +
                        "    RibbonCallbacks.OnRibbonLoad ribbon`r`n" +
                        "End Sub`r`n"
                $tpCodeModule.InsertLines($tpCodeModule.CountOfLines + 1, $code)
                Write-Host "Added ThisProject.OnRibbonLoad wrapper."
            }
            if ($tpCode -notmatch 'Public\s+Sub\s+GenerateDashboard\s*\(') {
                $code = "Public Sub GenerateDashboard(ByVal control As Object)`r`n" +
                        "    On Error Resume Next`r`n" +
                        "    RibbonCallbacks.GenerateDashboard control`r`n" +
                        "End Sub`r`n"
                $tpCodeModule.InsertLines($tpCodeModule.CountOfLines + 1, $code)
                Write-Host "Added ThisProject.GenerateDashboard wrapper."
            }
        } catch {
            Write-Warning ("Failed to add ThisProject wrappers: {0}" -f $_.Exception.Message)
        }
    } else {
        Write-Warning "ThisProject module not found; cannot add callback wrappers."
    }

    # --- PATCH D2: Fetch Ribbon XML
    Section "Loading RibbonX (customUI14) XML"
    $ribbonXml = Get-CustomUIXml -TemplateBaseWithRibbon $TemplateIn -FallbackXmlPath $CustomUiXmlFallback
    if (-not $ribbonXml) {
        Write-Warning "Ribbon XML not found in base or fallback: customUI14.xml"
    } elseif ($ribbonXml -notmatch '2009/07/customui') {
        Write-Warning "Ribbon XML does not appear to use customUI14 (2009/07) namespace."
    }

    # --- PATCH D2b: Apply RibbonX in-process (preferred)
    if (-not $SkipRibbonApply -and -not $UsePostSaveApply -and $ribbonXml) {
        Section "Applying RibbonX in-process (pre-save)"
        try {
            # Make UI visible to avoid SetCustomUI hang in headless mode
            $projApp.Visible = $true
            try { [System.Windows.Forms.Application]::DoEvents() | Out-Null } catch { }
            $projApp.ActiveProject.SetCustomUI($ribbonXml)
            Write-Host "RibbonX applied in-process (pre-save)."
        } catch {
            if ($AllowRibbonTimeout -or -not $RequireRibbon) {
                Write-Warning ("Pre-save RibbonX apply FAILED but continuing: {0}" -f $_.Exception.Message)
            } else {
                throw "Pre-save RibbonX apply FAILED: $($_.Exception.Message)"
            }
        }
    }

    # ---------- Save ----------
    Section "Saving output"
    Write-Host "Saving as: $TemplateOut"
    [void]$projApp.FileSaveAs($TemplateOut)
    if (-not (Test-Path -LiteralPath $TemplateOut)) {
        throw "Save failed: Output not found after save: $TemplateOut"
    }
    # Close and quit Project to release any lock before post-save Ribbon apply
    try { [void]$projApp.FileCloseAll() } catch { }
    try { $projApp.Quit() } catch { }
    if ($vbProj)  { Release-ComObject -ComObj $vbProj; $vbProj = $null }
    if ($projApp) { Release-ComObject -ComObj $projApp; $projApp = $null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    # --- PATCH D3: Optional post-save apply (fallback)
    Section "Applying RibbonX to output (post-save, STA + timeout)"
    if ($SkipRibbonApply) {
        Write-Warning "Skipping RibbonX apply step by request."
    } elseif (-not $UsePostSaveApply) {
        Write-Host "Post-save apply skipped (using in-process apply)."
    } elseif ($ribbonXml) {
        try {
            $attempt = 0
            $applied = $false
            while (-not $applied -and $attempt -le $RibbonRetries) {
                $attempt++
                Write-Host ("Ribbon apply attempt {0} (timeout: {1}s)" -f $attempt, $RibbonTimeoutSeconds)
                Apply-RibbonXToFileWithTimeout -FilePath $TemplateOut -RibbonXml $ribbonXml -TimeoutSeconds $RibbonTimeoutSeconds
                $applied = $true
            }
            Write-Host "RibbonX applied to output (post-save)."
        } catch {
            if ($AllowRibbonTimeout -or -not $RequireRibbon) {
                Write-Warning ("Post-build RibbonX apply FAILED but continuing: {0}" -f $_.Exception.Message)
            } else {
                throw "Post-build RibbonX apply FAILED: $($_.Exception.Message)"
            }
        }
    } else {
        if ($RequireRibbon) {
            throw "customUI14 is missing (no XML available to apply)."
        } else {
            Write-Warning "customUI14 is missing (no XML available to apply)."
        }
    }

    # ---------- Close ----------
    Section "Closing"
    Write-Host "Closing all documents..."
    try { [void]$projApp.FileCloseAll() } catch { }

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
    if ($projApp) {
        try { $projApp.Quit() } catch { }
    }
    if ($vbProj)  { Release-ComObject -ComObj $vbProj }
    if ($projApp) { Release-ComObject -ComObj $projApp }

    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    # Cleanup temp folders
    foreach ($tmp in @($TempImport, $TempImportNative)) {
        if ($tmp -and (Test-Path -LiteralPath $tmp)) {
            try { Remove-Item -LiteralPath $tmp -Recurse -Force } catch { }
        }
    }
}
