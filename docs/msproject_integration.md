# MS Project Integration — Windows & Linux (Wine) Strategy

## Key points
- MS Project is **Windows-only**. On Linux use **Wine** (limited) or a **Windows VM** (recommended).
- We **do not modify** business macros. We only **open projects** and **run existing macros**.
- All operations log to the VBE **Immediate Window** (`Ctrl+G`).

## Windows (development)
Requirements:
- Microsoft Project installed (Pro/Std).
- In Project: `File → Options → Trust Center → Trust Center Settings → Macro Settings` → enable macros for testing.

Usage:
- Import `msproject/ModuleMSProjectIntegration.bas` into your Excel `.xlsm`.
- Put your template at `templates/TemplateProject_v1.mpt`.
- In Excel VBE run `Test_OpenMpt` (smoke test) or `Test_OpenMptAndRunMacro` with a real macro name.

## Linux (targets)
### Wine
- Install Wine + winetricks; install MS Project into your Wine prefix.
- Use `tests/msproject/env_check.sh` to verify Wine prefix and path to `WINPROJ.EXE`.
- Headless COM on Linux isn’t supported; drive Project via Wine (or run automation on a Windows VM and trigger it remotely).

### Windows VM (recommended)
- Execute the COM automation inside the VM. Call it from CI/Linux via SSH, RDP, or a script runner.

## Macro execution model
- Run macro inside the opened project:
  `Application.Run "ProjectName!ModuleName.MacroName"`
- “Carrier project” option: open a second project that holds shared helper macros and run them against `ActiveProject` (no code injection into the business file).

## Logging
Examples:
