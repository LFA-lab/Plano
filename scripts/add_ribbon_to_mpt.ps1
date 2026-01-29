#Requires -Version 5.1
<#
.SYNOPSIS
  Inject a RibbonX (CustomUI14) into a Microsoft Project template (.mpt) on Windows PowerShell 5.1.

.DESCRIPTION
  - Zero manual steps.
  - Downloads OpenMcdf 2.3.0 (contains .NET Framework net40 build).
  - Compiles a tiny C# helper against OpenMcdf.dll (opens in Update mode and uses cf.Commit()).
  - Safely overwrites output with retry-based file unlock handling.
  - Reads input templates\TemplateBase.mpt (one folder above /scripts by default).
  - Writes output templates\TemplateBase_WithRibbon.mpt with the Ribbon embedded.
  - Uses CustomUI14 (2009/07) schema and verifies success.

.EXAMPLE
  .\scripts\add_ribbon_to_mpt.ps1
  .\scripts\add_ribbon_to_mpt.ps1 -Force -TabLabel "Plano" -OnAction "GenerateDashboard"
#>

[CmdletBinding()]
param(
    # ALIGNED DEFAULTS: read/write under /templates (sibling of /scripts)
    [string]$InputPath  = (Join-Path (Join-Path $PSScriptRoot "..\templates") "TemplateBase.mpt"),
    [string]$OutputPath = (Join-Path (Join-Path $PSScriptRoot "..\templates") "TemplateBase_WithRibbon.mpt"),

    [string]$TabLabel   = "Plano",
    [string]$OnAction   = "GenerateDashboard",
    [string]$OnLoad     = "OnRibbonLoad",

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Write-Info($msg)  { Write-Host "[INFO]  $msg" -ForegroundColor Cyan }
function Write-OK($msg)    { Write-Host "[OK]    $msg" -ForegroundColor Green }
function Write-Warn($msg)  { Write-Host "[WARN]  $msg" -ForegroundColor Yellow }
function Write-Err($msg)   { Write-Host "[ERROR] $msg" -ForegroundColor Red }

function Resolve-PathStrict {
    param([string]$Path)
    $full = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $full)) { throw "File not found: $full" }
    return $full
}

function Ensure-OutputWritable {
    param([string]$Path, [switch]$Force)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    if ((Test-Path -LiteralPath $Path) -and -not $Force) { throw "Output already exists: $Path. Use -Force to overwrite." }
}

# ---------------- File/dir helpers ----------------
function Test-IsDirectory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    return (Get-Item -LiteralPath $Path).PSIsContainer
}
function Test-FileLocked {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    if (Test-IsDirectory -Path $Path)        { return $false }
    try {
        $fs = [System.IO.File]::Open($Path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)
        $fs.Close(); return $false
    } catch { return $true }
}
function Wait-ForFileUnlock {
    param([string]$Path,[int]$TimeoutSeconds = 30,[int]$PollMs = 250)
    if (Test-IsDirectory -Path $Path) { return }
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $elapsed = 0
    while (Test-FileLocked -Path $Path) {
        Start-Sleep -Milliseconds $PollMs
        $elapsed += $PollMs
        if ($elapsed -ge ($TimeoutSeconds*1000)) { throw "Timed out waiting for file to be unlocked: $Path" }
    }
}
function Try-CloseProject {
    try {
        $procs = Get-Process -Name WINPROJ -ErrorAction SilentlyContinue
        if ($procs) {
            Write-Warn "Microsoft Project (WINPROJ.EXE) is running. Attempting to close..."
            $procs | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
    } catch { }
}

# ---------------- Ribbon XML + validation ----------------
function Get-RibbonXml {
    param([string]$TabLabel,[string]$OnAction,[string]$OnLoad)
@"
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="$OnLoad">
  <ribbon>
    <tabs>
      <tab id="tabCustom" label="$TabLabel">
        <group id="grpDashboard" label="Dashboard">
          <button id="btnGenerate"
                  label="Generate Dashboard"
                  size="large"
                  imageMso="Refresh"
                  onAction="$OnAction" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"@
}
function Test-CustomUiNamespace {
    param([string]$XmlText)
    if ($XmlText -notmatch 'http://schemas\.microsoft\.com/office/2009/07/customui') {
        throw "Ribbon XML missing CustomUI14 (2009/07) namespace."
    }
}
function Test-XmlWellFormed {
    param([string]$XmlText)
    try {
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver   = $null
        $sr = New-Object System.IO.StringReader($XmlText)
        $xr = [System.Xml.XmlReader]::Create($sr, $settings)
        while ($xr.Read()) { }
        $xr.Close(); $sr.Close()
    } catch {
        throw "Ribbon XML is not well-formed: $($_.Exception.Message)"
    }
}

# ------------- OpenMCDF loader (PS 5.1 / net40) -------------
$script:OpenMcdfAssembly = $null
$script:OpenMcdfDllPath  = $null

function Download-OpenMcdfLegacy {
    $runId      = [Guid]::NewGuid().ToString('N')
    $tempBase   = Join-Path $env:TEMP "OpenMcdf_$runId"
    $nupkgPath  = Join-Path $tempBase "OpenMcdf.2.3.0.nupkg"
    $extractDir = Join-Path $tempBase "pkg"

    New-Item -ItemType Directory -Path $tempBase   -Force | Out-Null
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    $url = "https://api.nuget.org/v3-flatcontainer/openmcdf/2.3.0/openmcdf.2.3.0.nupkg"
    Write-Info ("Downloading OpenMcdf 2.3.0 from NuGet: " + $url)
    Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $nupkgPath

    Write-Info ("Extracting package to " + $extractDir)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($nupkgPath, $extractDir)

    return $extractDir
}
function Get-OpenMcdfNet40Dll {
    param([string]$ExtractDir)
    $net40 = Join-Path $ExtractDir "lib\net40\OpenMcdf.dll"
    if (Test-Path -LiteralPath $net40) { return (Resolve-Path $net40).Path }
    $anyNet4 = Get-ChildItem -LiteralPath (Join-Path $ExtractDir "lib") -Recurse -Filter "OpenMcdf.dll" -ErrorAction SilentlyContinue |
               Where-Object { $_.FullName -match "\\lib\\net4" } | Select-Object -First 1
    if ($anyNet4) { return $anyNet4.FullName }
    $any = Get-ChildItem -LiteralPath $ExtractDir -Recurse -Filter "OpenMcdf.dll" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($any) { return $any.FullName }
    throw "OpenMcdf.dll not found inside downloaded package."
}
function Ensure-OpenMcdfLoaded {
    $localLib = Join-Path (Join-Path $PSScriptRoot "..") "lib\OpenMcdf.dll"
    $dllPath = $null

    if (Test-Path -LiteralPath $localLib) {
        $dllPath = (Resolve-Path $localLib).Path
        Write-Info ("Using local OpenMcdf: " + $dllPath)
    } else {
        $extractDir = Download-OpenMcdfLegacy
        $dllPath    = Get-OpenMcdfNet40Dll -ExtractDir $extractDir
        Write-Info ("Using downloaded OpenMcdf: " + $dllPath)
    }

    $script:OpenMcdfAssembly = [Reflection.Assembly]::LoadFrom($dllPath)
    if (-not $script:OpenMcdfAssembly) { throw "OpenMcdf assembly did not load from $dllPath." }

    $script:OpenMcdfDllPath = $dllPath
    Write-OK ("OpenMcdf loaded: " + $script:OpenMcdfAssembly.FullName)
}

# ---------------- C# helper (use single-quoted here-string) ----------------
function Ensure-CSharpInjector {
    $typeExists = [Type]::GetType('MptRuntime.MptRibbonWriter')
    if ($typeExists) { return }
    if (-not $script:OpenMcdfDllPath) { throw "OpenMcdf not prepared." }

$src = @'
using System;
using OpenMcdf;

namespace MptRuntime
{
    public static class MptRibbonWriter
    {
        public static void Inject(string path, byte[] data)
        {
            using (var cf = new CompoundFile(path, CFSUpdateMode.Update, CFSConfiguration.Default))
            {
                var root = cf.RootStorage;

                try { root.Delete("customUI14"); } catch {}

                var s = root.AddStream("customUI14");
                s.SetData(data);

                cf.Commit();
            }
        }

        public static bool Verify(string path)
        {
            using (var cf = new CompoundFile(path, CFSUpdateMode.ReadOnly, CFSConfiguration.Default))
            {
                try
                {
                    var s = cf.RootStorage.GetStream("customUI14");
                    var bytes = s.GetData();
                    return bytes != null && bytes.Length > 0;
                }
                catch
                {
                    return false;
                }
            }
        }
    }
}
'@

    Add-Type -TypeDefinition $src -Language CSharp -ReferencedAssemblies $script:OpenMcdfDllPath -PassThru | Out-Null
}

# ------------- Inject / Verify wrappers -------------
function Inject-CustomUI14 {
    param([string]$TargetPath,[string]$XmlText)
    Test-CustomUiNamespace -XmlText $XmlText
    Test-XmlWellFormed     -XmlText $XmlText
    Wait-ForFileUnlock -Path $TargetPath -TimeoutSeconds 30

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($XmlText)
    [MptRuntime.MptRibbonWriter]::Inject($TargetPath, $bytes)
}

function Verify-CustomUI14 {
    param([string]$TargetPath)
    Wait-ForFileUnlock -Path $TargetPath -TimeoutSeconds 30
    if (-not [MptRuntime.MptRibbonWriter]::Verify($TargetPath)) {
        throw "customUI14 stream missing or empty."
    }
    return $true
}

# --------------------- MAIN ---------------------
try {
    Write-Info ("Input:  " + $InputPath)
    Write-Info ("Output: " + $OutputPath)

    Try-CloseProject

    $inFull  = Resolve-PathStrict -Path $InputPath
    Ensure-OutputWritable -Path $OutputPath -Force:$Force

    Ensure-OpenMcdfLoaded
    Ensure-CSharpInjector

    $xml = Get-RibbonXml -TabLabel $TabLabel -OnAction $OnAction -OnLoad $OnLoad

    # Cleanly overwrite output
    if (Test-Path -LiteralPath $OutputPath) {
        if ($Force) {
            Write-Info "Removing existing output (Force): $OutputPath"
            try { Remove-Item -LiteralPath $OutputPath -Force -ErrorAction Stop } catch {}
            if (Test-Path -LiteralPath $OutputPath) { Wait-ForFileUnlock -Path $OutputPath -TimeoutSeconds 30 }
        } else {
            throw "Output already exists: $OutputPath. Use -Force to overwrite."
        }
    }

    # Copy base template to output
    Copy-Item -LiteralPath $inFull -Destination $OutputPath -Force
    Write-OK "Copied base template to output."

    # Extra wait before opening via OpenMCDF (AV/scanner may touch it right after copy)
    Wait-ForFileUnlock -Path $OutputPath -TimeoutSeconds 30

    Inject-CustomUI14 -TargetPath $OutputPath -XmlText $xml
    Verify-CustomUI14 -TargetPath $OutputPath | Out-Null
    Write-OK "Ribbon (customUI14) embedded successfully."

    # Optionally open the output template in Microsoft Project (kept as-is)
    try {
        Write-Info "Opening output template in Microsoft Project..."
        Start-Process -FilePath $OutputPath | Out-Null
    } catch {
        Write-Warn "Could not automatically open the output file: $($_.Exception.Message)"
    }

    Write-Host ""
    Write-OK ("Done. Output ready: " + $OutputPath)
    Write-Host ("Open the .mpt in Microsoft Project to see tab '" + $TabLabel + "' with a 'Generate Dashboard' button.") -ForegroundColor Gray
    exit 0
}
catch {
    Write-Err $_.Exception.Message
    if ($_.Exception.InnerException) { Write-Err $_.Exception.InnerException.Message }
    Write-Host "Tip: If this is a file lock issue, ensure Microsoft Project is closed and no AV/sync tool is holding the file. You can also try: Stop-Process -Name WINPROJ -Force" -ForegroundColor DarkGray
    exit 1
}