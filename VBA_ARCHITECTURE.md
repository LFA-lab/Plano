# Plano VBA Architecture - Developer Guide

## Quick Start

### Building the Template

```powershell
cd scripts
.\build_mpt.ps1
```

This creates `templates/ModeleImport.mpt` from `TemplateBase_WithRibbon.mpt` by injecting all VBA modules from `macros/production/`.

### VBA Module Structure

```
macros/production/
├── ThisProject.cls          ← Event handlers (auto-show UserForm/.mpt, menu/.mpp)
├── modPlanoMenu.bas          ← Menu creation via CommandBars
├── PlanoMenuActions.bas      ← Menu action handlers (GenerateDashboard, etc.)
├── Import_OPTIMISE.vb        ← Excel → MS Project import engine
├── PlanoCore.bas             ← Core utilities (file selection, etc.)
├── ExportToJson.bas          ← JSON export for dashboard
└── generatevb.bas            ← VBA code generation utilities
```

## Architecture Overview

### Workflow

```
┌─────────────────────────────────────────────────────────────────┐
│ 1. User Opens ModeleImport.mpt                                  │
│    ↓                                                             │
│    ThisProject.Project_Open() detects .mpt extension            │
│    ↓                                                             │
│    UserFormImport.Show (automatic)                              │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│ 2. User Selects Excel File via UserFormImport                   │
│    ↓                                                             │
│    btnBrowseFile_Click() → Excel file picker                    │
│    ↓                                                             │
│    ImportDataSilent(filePath)                                   │
│    ↓                                                             │
│    Import_Taches_Simples_AvecTitre_WithFile(filePath)          │
│    ↓                                                             │
│    Import_OPTIMISE creates tasks from Excel                     │
│    ↓                                                             │
│    FileSaveAs → .mpp file saved next to Excel                  │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│ 3. User Opens .mpp File                                          │
│    ↓                                                             │
│    ThisProject.Project_Open() detects .mpp extension            │
│    ↓                                                             │
│    CreatePlanoMenu() → Adds "Plano" to menu bar                │
│    ↓                                                             │
│    User clicks Plano → menu items call PlanoMenuActions        │
└─────────────────────────────────────────────────────────────────┘
```

### Key Design Principles

#### 1. **No Ribbon Code** ❌
- Architecture uses **CommandBars** only (not Ribbon customUI)
- Compatible with all MS Project versions (2016+)
- No OpenMCDF, no customUI14.xml, no PowerShell Ribbon injection

#### 2. **Portable Paths** ✅
- ALWAYS use: `Environ$("USERPROFILE")`, `Application.TemplatesPath`, `ActiveProject.Path`
- NEVER use: `"C:\Users\..."`, `"D:\..."`, hardcoded usernames

#### 3. **Relative OnAction** ✅
- Menu items: `btn.OnAction = "MacroName"` (relative)
- NEVER: `btn.OnAction = "File.mpt!MacroName"` (absolute)

#### 4. **Event-Driven Workflow** 🔄
- `Project_Open()` auto-detects .mpt vs .mpp
- `Project_BeforeClose()` cleans up menu
- No manual initialization required

## Module Reference

### ThisProject.cls

**Purpose:** Event handlers for automatic workflow

**Key Methods:**
```vba
Private Sub Project_Open(ByVal pj As Project)
    ' Detects file extension and:
    ' .mpt → Show UserFormImport
    ' .mpp → Create Plano menu

Private Sub Project_BeforeClose(ByVal pj As Project)
    ' Clean up Plano menu
```

**Critical:** Must be included in build, cannot be removed

---

### modPlanoMenu.bas

**Purpose:** Menu creation and lifecycle management

**Key Methods:**
```vba
Public Sub CreatePlanoMenu()
    ' Creates "Plano" menu in CommandBars
    ' Tries: "Menu Bar" → "Menu Commands" → "Ribbon"

Public Sub RemovePlanoMenu()
    ' Removes "Plano" menu from all CommandBars

Private Sub AddPlanoButton(parent, caption, macroName, faceId)
    ' Helper: Adds menu item with OnAction = macroName
```

**Menu Items Created:**
- Generate Dashboard → `GenerateDashboard`
- Import from Excel → `ImportFromExcel`
- Export to JSON → `ExportData`
- Control Panel → `ShowPlanoControl`

---

### PlanoMenuActions.bas

**Purpose:** Implementation of all menu action handlers

**Public Subs (called by menu OnAction):**
```vba
Public Sub GenerateDashboard()
    ' Exports project data, generates dashboard

Public Sub ImportFromExcel()
    ' Shows file picker, imports Excel data

Public Sub ExportData()
    ' Exports project to JSON

Public Sub ShowPlanoControl()
    ' Shows control panel or info dialog
```

**Important:** All functions must be `Public Sub` with no parameters (required for OnAction)

---

### Import_OPTIMISE.vb

**Purpose:** Excel → MS Project import engine

**Main Function:**
```vba
Sub Import_Taches_Simples_AvecTitre()
    ' Shows file picker, imports Excel, creates tasks

Sub Import_Taches_Simples_AvecTitre_WithFile(filePath As String)
    ' Wrapper that accepts pre-selected file (from UserFormImport)
    ' Avoids double file selection dialog
```

**Workflow:**
1. Read Excel file (all data into memory array for performance)
2. Create MS Project structure (tasks, resources, assignments)
3. Apply tags (Zone, Sous-Zone, Tranche, Type, Entreprise, Niveau, Onduleur, PTR)
4. Save log file next to Excel

**Performance Optimizations:**
- Array-based Excel reading (not cell-by-cell)
- Resource caching via Dictionary
- ScreenUpdating disabled during import

---

### PlanoCore.bas

**Purpose:** Core utilities and helper functions

**Public Functions:**
```vba
Public Function DownloadsFolder() As String
    ' Returns portable path to user's Downloads folder

Public Sub CreateProjectFromTemplate(templatePath, outputPath)
    ' Creates new project from template

Public Function SelectPlanningFile() As String
    ' Shows file picker for Excel/MPP files

Public Sub ImportDataSilent(filePath As String)
    ' Imports file based on extension (.mpp, .xlsx, .xlsm, .csv)
```

---

### ExportToJson.bas

**Purpose:** Export project data to JSON for HTML dashboard

**Main Function:**
```vba
Public Sub ExportToJson()
    ' Exports tasks, resources, assignments to JSON
    ' Used by dashboard visualization
```

---

## Common Tasks

### Adding a New Menu Item

1. **Define constant in modPlanoMenu.bas:**
   ```vba
   Public Const MACRO_MYFEATURE As String = "MyFeature"
   ```

2. **Add button in CreatePlanoMenu():**
   ```vba
   AddPlanoButton pop, "My Feature", MACRO_MYFEATURE, 123
   ```

3. **Implement handler in PlanoMenuActions.bas:**
   ```vba
   Public Sub MyFeature()
       MsgBox "My feature!", vbInformation
   End Sub
   ```

4. **Rebuild template:**
   ```powershell
   .\scripts\build_mpt.ps1
   ```

---

### Debugging the Workflow

#### Issue: UserFormImport doesn't show automatically
- Check: Is ThisProject.cls included in build?
- Check: Does Project_Open() call UserFormImport.Show?
- Test: Add `MsgBox "Open: " & fileExt` in Project_Open()

#### Issue: Plano menu doesn't appear
- Check: Is .mpp extension detected correctly?
- Check: Does CreatePlanoMenu() complete without errors?
- Test: Call `modPlanoMenu.CreatePlanoMenu` manually in Immediate Window

#### Issue: Menu items show errors
- Check: Does the Public Sub exist in PlanoMenuActions.bas?
- Check: Is function name spelled correctly (case-sensitive)?
- Check: Is PlanoMenuActions.bas included in build?

---

## Testing

### Manual Test Checklist

**Template Test (.mpt):**
1. Open `ModeleImport.mpt`
2. ✅ UserFormImport appears automatically
3. ✅ Select Excel file (one dialog only)
4. ✅ Import creates tasks correctly
5. ✅ .mpp saved next to Excel
6. ✅ No Plano menu in .mpt

**Project Test (.mpp):**
1. Open generated .mpp
2. ✅ Plano menu appears automatically
3. ✅ No UserFormImport in .mpp
4. ✅ Menu items work (Dashboard, Export)
5. Close .mpp
6. ✅ Menu cleaned up

---

## Troubleshooting

### Build Failures

**Error: "Cannot access the VBA project"**
- Fix: Enable "Trust access to the VBA project object model" in MS Project Trust Center

**Error: "No macros imported (X expected)"**
- Check: Are VBA files in `/macros/production/`?
- Check: Do .bas files have `Attribute VB_Name = "..."`?
- Run: `.\scripts\build_mpt.ps1 -Verbose` for details

### Runtime Errors

**Error: "Sub or Function not defined" (menu click)**
- Check: Is PlanoMenuActions.bas included in build?
- Check: Does function name match exactly? (case-sensitive)
- Fix: Rebuild template

**Error: "Automation error" (GenerateDashboard)**
- Check: Does ExportToJsonModule exist?
- Check: Is Excel installed?
- Fix: Add error handling in PlanoMenuActions.bas

**Error: Double file selection dialogs**
- Check: Is Import_Taches_Simples_AvecTitre_WithFile called?
- Check: Are g_UsePreSelectedFile flags set correctly?
- Fix: Verify UserFormImport calls wrapper function

---

## Architecture Validation

### ✅ Checklist for Code Reviews

- [ ] No hardcoded paths (C:\, D:\, user names)
- [ ] All paths use Environ$() or Application.TemplatesPath
- [ ] Menu OnAction uses relative macro names only
- [ ] No Ribbon code (OnRibbonLoad, customUI, etc.)
- [ ] ThisProject.cls handles both .mpt and .mpp
- [ ] All Public Subs in PlanoMenuActions have no parameters
- [ ] Error handling present in all menu actions
- [ ] Build script includes all required modules

---

## References

- **Full Audit Report:** `AUDIT_REPORT.md`
- **Implementation Details:** `IMPLEMENTATION_SUMMARY.md`
- **Build Script:** `scripts/build_mpt.ps1`
- **Main README:** `README.md`

---

## Support

For questions or issues with the VBA architecture:
1. Check `AUDIT_REPORT.md` for detailed analysis
2. Check `IMPLEMENTATION_SUMMARY.md` for implementation notes
3. Run build with `-Verbose` flag for diagnostic output
4. Open issue on GitHub: https://github.com/LFA-lab/Plano/issues

---

**Last Updated:** 2026-02-15
**Architecture Version:** 1.0 (CommandBars-based)
