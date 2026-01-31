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
The build_mpt.ps1 script imports these `.bas` files into the Microsoft Project Template.

## How to Modify a Macro
1. Open the `.bas` file inside `/macros/production/`.
2. Edit using any text editor (VS Code recommended or VBA editor).
3. Save your changes.
4. Do NOT edit macros directly inside the .mpt file—those changes will be overwritten.

## Required Tools
- Git
- PowerShell 5+
- Microsoft Project

## Developer Workflow
### 1. Modify the Macro
Edit the relevant `.bas` file in `/macros/production/`.

### 2. Stage and Commit Changes
```
git add .
git commit -m "Updated macro XYZ"
```

### 3. Run the Push Script
```
./scripts/push.ps1
```
Or:
```
./scripts/push.ps1 -NoAmend
```

### 4. What the Script Does Automatically
- Runs build_mpt.ps1 → generates a fresh template
- Runs add_ribbon_to_mpt.ps1 → injects Ribbon UI
- Writes final file to `/templates/ModèleImport.mpt`
- Creates a Git commit automatically
- Amends the previous commit unless `-NoAmend` is used
- Performs a safe `git push`

### Clarifying Amend Behavior
- Default: amends last commit (keeps history clean)
- Use -NoAmend to create a new commit

## What NOT to Do
- Do NOT push using git push directly → always use push.ps1.
- Do NOT manually modify `templates/ModèleImport.mpt`.
- Do NOT edit macros inside Project.
- Do NOT change Ribbon XML manually.

## Troubleshooting
### Build Fails
- Invalid VBA syntax
- Microsoft Project missing
- File locked

### Template Not Updated
```
pwsh ./scripts/build_mpt.ps1
```
Check output in `/templates/ModèleImport.mpt`.

### Git Errors
```
git pull --rebase
```
Resolve conflicts → rerun scripts.
