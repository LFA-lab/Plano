# WORKFLOW_DEV.md

## Purpose
This document explains how developers modify VBA macros, rebuild the Microsoft Project template, and safely push updates using the automation scripts.

## Macro Storage Structure
Macros are stored inside the repository here:
```
/macros/production/
    Module1.bas
    Module2.bas
    ...
```
The `build_mpt.ps1` script imports these `.bas` files into the Microsoft Project template build.

## How to Modify a Macro
1. Open the `.bas` file inside `/macros/production/`.
2. Edit using any text editor (VS Code recommended or the VBA editor).
3. Save your changes.
4. **Do NOT edit macros directly inside the .mpt file**—those changes will be overwritten by the build.

## Required Tools
- Git
- PowerShell 5+
- Microsoft Project

## Developer Workflow
Follow this workflow whenever modifying VBA code.

### 1) Modify the Macro
Edit the relevant `.bas` file in `/macros/production/`.

### 2) Stage and Commit Changes
```powershell
git add .
git commit -m "Updated macro XYZ"
```

### 3) Run the Push Script
From the repository root:
```powershell
./scripts/push.ps1
```
If you **don’t** want to amend the previous commit:
```powershell
./scripts/push.ps1 -NoAmend
```

### 4) What Happens Automatically
`push.ps1` orchestrates the full release pipeline:
- Runs `build_mpt.ps1` → regenerates the base template as `templates/TemplateBase.mpt`
- Runs `add_ribbon_to_mpt.ps1` → injects the custom Ribbon and outputs `templates/TemplateBase_WithRibbon.mpt`
- Delegates all Git logic to `scripts/commit_and_push.ps1`
  - **Default:** amend previous commit
  - **With `-NoAmend`:** create a new commit
- Publishes the final artifact as `templates/ModèleImport.mpt`

#### Console Output (What you should see)
```
🎨 Step 1: Injecting ribbon...
✅ Ribbon injected successfully.
🔨 Step 2: Building ModèleImport.mpt...
✅ Build successful. Macros imported: X/X
📦 Pushing to GitHub...
✅ Push successful.
```

### Clarifying “Amend” Behavior
- **Default (`push.ps1`)**: Amends the last commit so build artifacts don’t create extra commits.
- **`-NoAmend`**: `push.ps1` forwards the switch to `commit_and_push.ps1`, which creates a new standalone commit.

## ❌ What NOT to Do
- Do **NOT** push using `git push` directly → always use `push.ps1` so the build+UI+commit flow stays consistent.
- Do **NOT** manually modify any files in `templates/`.
- Do **NOT** edit macros inside Microsoft Project.
- Do **NOT** change the Ribbon XML manually.

## Troubleshooting

### Build Fails
Common causes:
- Invalid VBA syntax in `/macros/production/*.bas`
- Microsoft Project not installed
- File locked/in use

**Fix:** Review PowerShell output → correct the error → re-run the script.

### Template Not Updated
- Script not executed from repository root
- Permissions to write to `/templates/` missing

**Fix:**
```powershell
pwsh ./scripts/build_mpt.ps1
```
Check the outputs at:
```
/templates/TemplateBase.mpt
/templates/TemplateBase_WithRibbon.mpt
/templates/ModèleImport.mpt
```

### Git Errors
Most Git-related errors are surfaced by `scripts/commit_and_push.ps1` (called by `push.ps1`). Typical issues:
- Remote ahead / rebase needed
- Merge conflicts
- Unstaged local changes

**Fix:**
```powershell
git pull --rebase
```
Resolve conflicts →
```powershell
./scripts/commit_and_push.ps1  # or simply rerun ./scripts/push.ps1
```
