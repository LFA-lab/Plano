#Requires -Version 5.1
<#
.SYNOPSIS
  Rebuilds the Project template (Ribbon + macros) and pushes code + .mpt together.

.DESCRIPTION
  - Executes scripts\add_ribbon_to_mpt.ps1 first (injects ribbon into base template).
  - Then executes scripts\build_mpt.ps1 (reads templates\TemplateBase_WithRibbon.mpt and writes templates\ModèleImport.mpt).
  - If build succeeds: adds templates\ModèleImport.mpt to Git, amends the last commit by default, then pushes.
  - If build fails: prints error and exits without pushing.
  - Works from anywhere; resolves repo root from this script's location.
  - DRY RUN mode: runs build, skips all Git checks and operations; prints intended actions.

.USAGE
  PS> .\scripts\push.ps1
  PS> .\scripts\push.ps1 -Verbose
  PS> .\scripts\push.ps1 -NoAmend           # create a new commit instead of amending the previous one
  PS> .\scripts\push.ps1 -DryRun -Verbose   # simulate everything except Git checks and writes

.NOTES
  - Requires Microsoft Project (for the build).
  - If your console shows `ModÃ¨leImport.mpt`, it is a display issue only; the file path is correct.
#>

[CmdletBinding()]
param(
    [switch]$NoAmend,
    [switch]$DryRun
)

# Make console UTF-8 friendly for display
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Global dry-run flag
$script:DryRunMode = $DryRun.IsPresent
if ($script:DryRunMode) {
    Write-Host "=== DRY RUN: build will execute; all Git checks and operations are disabled ===" -ForegroundColor Yellow
}

function Section([string]$Text) {
    Write-Host ""
    Write-Host "================================="
    Write-Host $Text
    Write-Host "================================="
}

# Resolve Git executable path reliably on Windows
$script:GitExe = $null
function Resolve-GitExe {
    if ($script:GitExe) { return $script:GitExe }

    try {
        $paths = & where.exe git 2>$null
        if ($LASTEXITCODE -eq 0 -and $paths) {
            $cands = $paths -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            # Prefer cmd shim
            $pref = $cands | Where-Object { $_ -match '\\Git\\cmd\\git\.exe$' } | Select-Object -First 1
            if ($pref) { $script:GitExe = $pref; return $script:GitExe }
            # Else prefer bin
            $pref = $cands | Where-Object { $_ -match '\\Git\\bin\\git\.exe$' } | Select-Object -First 1
            if ($pref) { $script:GitExe = $pref; return $script:GitExe }
            # Else first hit
            $script:GitExe = $cands[0]
            return $script:GitExe
        }
    } catch { }

    # Fallback locations
    $common = @(
        'C:\Program Files\Git\cmd\git.exe',
        'C:\Program Files\Git\bin\git.exe',
        'C:\Program Files (x86)\Git\cmd\git.exe',
        'C:\Program Files (x86)\Git\bin\git.exe'
    )
    foreach ($p in $common) {
        if (Test-Path -LiteralPath $p) { $script:GitExe = $p; return $script:GitExe }
    }

    throw "Git executable not found via PATH or common locations. Ensure Git is installed and on PATH."
}

# PS5.1-safe Git invocation; accepts argument array and runs inside repo root
function Invoke-Git {
    param(
        [Parameter(Mandatory)][string[]]$Args,
        [switch]$IgnoreErrors
    )

    if ($script:DryRunMode) {
        Write-Host ("[DRYRUN] git {0}" -f ($Args -join ' '))
        return @{ Code=0; Out=""; Err="" }
    }

    $gitExe = Resolve-GitExe
    Write-Verbose ("{0} {1}" -f $gitExe, ($Args -join ' '))

    $ErrorActionPreference_ = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $output = & $gitExe @Args 2>&1
    $code = $LASTEXITCODE
    $ErrorActionPreference = $ErrorActionPreference_

    $stdout = ($output | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
    if (-not $IgnoreErrors -and $code -ne 0) {
        throw ("git {0} failed ({1}):`n{2}" -f ($Args -join ' '), $code, $stdout.Trim())
    }
    return @{ Code=$code; Out=$stdout; Err="" }
}

# PS5.1-safe relative path
function Get-RelativePath {
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$TargetPath
    )
    $base = (Resolve-Path -LiteralPath $BasePath).Path
    $tgt  = (Resolve-Path -LiteralPath $TargetPath).Path
    if (-not $base.EndsWith('\')) { $base = $base + '\' }
    $baseUri = New-Object System.Uri($base)
    $tgtUri  = New-Object System.Uri($tgt)
    $rel = $baseUri.MakeRelativeUri($tgtUri).ToString()
    $rel = [System.Uri]::UnescapeDataString($rel) -replace '/', '/'
    return $rel
}

# ---------- Resolve repo root and important paths ----------
$ScriptDir   = Split-Path -Parent $PSCommandPath
$RepoRoot    = (Get-Item $ScriptDir).Parent.FullName
$BuildScript = Join-Path $RepoRoot "scripts\build_mpt.ps1"
$RibbonScript= Join-Path $RepoRoot "scripts\add_ribbon_to_mpt.ps1"
$MptPath     = Join-Path $RepoRoot "templates\ModèleImport.mpt"  # accent by design

Section 'Pre-flight'

# Check script presence
if (-not (Test-Path -LiteralPath $BuildScript)) {
    throw "Build script not found: $BuildScript"
}
if (-not (Test-Path -LiteralPath $RibbonScript)) {
    throw "Ribbon injection script not found: $RibbonScript"
}

# Ensure all git commands run from repo root
$oldLocation = Get-Location
Push-Location -LiteralPath $RepoRoot
$popNeeded = $true

try {
    # Git checks: SKIPPED entirely in DryRun
    if ($script:DryRunMode) {
        Write-Verbose "Dry run: skipping Git availability and repository checks."
        $currentBranch = '<dryrun-branch>'
    } else {
        # Check Git availability
        try {
            $gitVer = Invoke-Git @("--version")
            Write-Verbose ($gitVer.Out.Trim())
        } catch {
            throw "Git not found. Please install Git and ensure it is on PATH."
        }

        # Ensure we are inside a git work tree
        $inside = Invoke-Git @("rev-parse", "--is-inside-work-tree")
        if ($inside.Out.Trim() -ne "true") {
            throw "Not inside a Git repository. Run this script from a cloned repo."
        }

        # Resolve current branch
        $currentBranch = (Invoke-Git @("rev-parse", "--abbrev-ref", "HEAD")).Out.Trim()
        if ([string]::IsNullOrWhiteSpace($currentBranch) -or $currentBranch -eq "HEAD") {
            throw "Cannot determine current branch (detached HEAD?)."
        }
    }

    Write-Host ("Current branch: {0}" -f $currentBranch)

    # ---------- NEW: Inject ribbon first ----------
    Write-Host "Step 1: Injecting ribbon..." -ForegroundColor Cyan
    powershell.exe -ExecutionPolicy Bypass -File "$RepoRoot\scripts\add_ribbon_to_mpt.ps1"

    if ($LASTEXITCODE -ne 0) {
        Write-Host "[FAIL] Ribbon injection failed. Push aborted." -ForegroundColor Red
        exit 1
    }

    Write-Host "[OK] Ribbon injected successfully." -ForegroundColor Green

    # ---------- Run build ----------
    Section 'Building template via build_mpt.ps1'
    Write-Host 'Step 2: Building ModèleImport.mpt...' -ForegroundColor Cyan
    try {
        $psArgs = @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass',
            '-File', $BuildScript
        )
        Write-Verbose ('powershell.exe {0}' -f ($psArgs -join ' '))
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $psArgs -Wait -PassThru
        if ($p.ExitCode -ne 0) {
            throw ('build_mpt.ps1 failed with exit code {0}' -f $p.ExitCode)
        }
    }
    catch {
        Write-Host ''
        Write-Host 'Build failed. Not pushing changes.' -ForegroundColor Red
        $msg = if ($_.Exception) { $_.Exception.Message } else { $_.ToString() }
        Write-Host ('Reason: {0}' -f $msg) -ForegroundColor Yellow
        Write-Host ''
        Write-Verbose ('Full error record:' + [Environment]::NewLine + ('{0}' -f $_))
        exit 1
    }

    # Verify output exists
    if (-not (Test-Path -LiteralPath $MptPath)) {
        Write-Error ('Build completed but output not found: {0}' -f $MptPath)
        Write-Host 'Not pushing changes.' -ForegroundColor Yellow
        exit 1
    }

    # You can replace X/X below by parsing build logs if needed.
    Write-Host 'Build successful. Macros imported: X/X'
    Write-Host ('Build OK. Output: {0}' -f $MptPath)

    # Compute relative path for git
    $relativeMpt = Get-RelativePath -BasePath $RepoRoot -TargetPath $MptPath
    if ([string]::IsNullOrWhiteSpace($relativeMpt)) { $relativeMpt = 'templates/ModèleImport.mpt' }
    $relativeMpt = $relativeMpt -replace '\\','/'

    # ---------- Stage the .mpt ----------
    Section 'Staging built template'
    if ($script:DryRunMode) {
        Write-Host ('[DRYRUN] Would stage: {0}' -f $relativeMpt)
    } else {
        Invoke-Git @('add', '--', $relativeMpt)
    }

    # Check staged changes
    $diffIndex = if ($script:DryRunMode) {
        Write-Host '[DRYRUN] Would check staged changes.'
        ''
    } else {
        (Invoke-Git @('diff', '--cached', '--name-only')).Out.Trim()
    }

    # Prepare flag for commit logic
    $hasChanges = -not [string]::IsNullOrWhiteSpace($diffIndex)

    # ---------- Commit & Push with user-friendly error handling ----------
    Section 'Committing & Pushing'
    $commitPushScript = Join-Path $RepoRoot 'scripts\commit_and_push.ps1'
    . $commitPushScript

    Section 'Done'
    if ($script:DryRunMode) {
        Write-Host 'Dry run complete. No changes were pushed.'
    } else {
        Write-Host 'Push complete.'
    }
}
finally {
    if ($popNeeded) { Pop-Location | Out-Null }
}