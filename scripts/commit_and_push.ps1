# Commit and push logic - dot-sourced from push.ps1 to avoid nested try/catch parse issues.
# Expects in scope: $hasChanges, $currentBranch, $NoAmend, and script:DryRunMode; uses Invoke-Git.
trap {
    Write-Host ""
    Write-Host "[FAIL] Git operation failed." -ForegroundColor Red
    Write-Host ""
    Write-Host "Common causes:" -ForegroundColor Yellow
    Write-Host " • Authentication failed (check GitHub token or SSH key)"
    Write-Host " • Merge conflict (pull first and resolve conflicts)"
    Write-Host " • Remote not configured (run: git remote -v)"
    Write-Host " • Network issue (check connection)"
    Write-Host ""
    Write-Host ("Error details: {0}" -f $_) -ForegroundColor Gray
    exit 1
}
if ($hasChanges) {
    if ($NoAmend) {
        $msg = "Add ModèleImport.mpt (auto build)"
        if ($script:DryRunMode) { Write-Host ('[DRYRUN] Would commit with message: {0}' -f $msg) }
        else { Invoke-Git @("commit", "-m", $msg); Write-Host "Created a new commit for the updated template." }
    } else {
        if ($script:DryRunMode) { Write-Host '[DRYRUN] Would amend the last commit to include the built template.' }
        else {
            $hasHead = (Invoke-Git @("rev-parse", "--verify", "HEAD") -IgnoreErrors).Code -eq 0
            if ($hasHead) { Invoke-Git @("commit", "--amend", "--no-edit"); Write-Host "Amended the last commit to include the built template." }
            else { $msg = "Initial commit with ModèleImport.mpt (auto build)"; Invoke-Git @("commit", "-m", $msg); Write-Host "Created initial commit." }
        }
    }
} else { Write-Host "No staged changes (template unchanged). Skipping commit step." }

if (-not $script:DryRunMode) {
    $remotesRaw = (Invoke-Git @("remote")).Out
    $remotes = $remotesRaw -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if ($remotes -notcontains 'origin') {
        Write-Host "[FAIL] No 'origin' remote configured." -ForegroundColor Red
        Write-Host 'Add remote with: git remote add origin <repository-url>' -ForegroundColor Yellow
        exit 1
    }
} elseif ($script:DryRunMode) { Write-Host '[DRYRUN] Would validate that ''origin'' remote is configured.' }

if ($script:DryRunMode) { Write-Host '[DRYRUN] Would push to upstream.' }
else {
    Write-Host "Pushing to GitHub..." -ForegroundColor Cyan
    $upstreamRef = '@{u}'
    $hasUpstream = (Invoke-Git @("rev-parse", "--abbrev-ref", "--symbolic-full-name", $upstreamRef) -IgnoreErrors).Code -eq 0
    if ($hasUpstream) { Invoke-Git @("push") }
    else { Write-Host ("No upstream configured. Setting upstream to origin/{0} ..." -f $currentBranch); Invoke-Git @("push", "-u", "origin", $currentBranch) }
    Write-Host "[OK] Push successful." -ForegroundColor Green
}
