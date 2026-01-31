# ARCHITECTURE.md

## Overview
This document provides a **full technical explanation** of the local automation that builds and publishes a Microsoft Project template. The system is intentionally **simple, script‑driven, and CI‑free**, so developers can build and push from their workstation while consultants consume a single, signed‑off template file.

## Key Objectives
- **Consistency:** One reproducible pipeline produces one canonical output: `/templates/ModèleImport.mpt`.
- **Safety:** Macros are stored as text under `/macros/production/`, never inside binary Project files.
- **Simplicity:** A single command (`./scripts/push.ps1`) runs the entire build and publication workflow.

## Repository Structure
```
Repository Root
 ├── macros/
 │     └── production/          # VBA source files (.bas)
 ├── scripts/
 │     ├── build_mpt.ps1
 │     ├── add_ribbon_to_mpt.ps1
 │     └── push.ps1
 ├── templates/
 │     └── ModèleImport.mpt
 └── docs/
```

## Environments & Prerequisites
- **Windows + PowerShell 5+**
- **Microsoft Project installed** (required for build automation)
- **Git installed and configured**
- **UTF‑8 enabled shell** (due to accented filename `ModèleImport.mpt`)

## High-Level Data Flow
```
/macros/production/*.bas
    → build_mpt.ps1
    → intermediate template
    → add_ribbon_to_mpt.ps1
    → templates/ModèleImport.mpt
    → push.ps1 → Git remote
```

## Script Responsibilities

### build_mpt.ps1 — Template Assembly
- Opens minimal Project template
- Purges old modules
- Imports `.bas` modules from `/macros/production/`
- Saves template as `/templates/ModèleImport.mpt`
- Guarantees **idempotency** and **deterministic builds**

### add_ribbon_to_mpt.ps1 — Ribbon Injection
- Opens `/templates/ModèleImport.mpt`
- Replaces (not appends) the Ribbon XML definition
- Ensures callbacks match existing macros

### push.ps1 — Orchestration & Git Workflow
- Runs both build scripts
- Stages artifacts
- Creates a commit
  - Default: **amends** previous commit
  - With `-NoAmend`: creates a separate new commit
- Performs safe push to remote

## Why This Architecture Works
- **Idempotent:** Same inputs always produce same template.
- **Versionable:** Source is stored in Git as text.
- **Zero Drift:** Ribbon XML is always replaced, never appended.
- **No CI dependency:** All operations are local and transparent.

## Deep-Dive: Reproducibility
- Modules imported alphabetically ensure stable ordering
- Ribbon replaced fully to avoid partial mismatches
- Final `.mpt` generated only via scripts (no manual edits)

## Security Considerations
- Macro warnings are expected; macros must be enabled
- Optional enterprise signing available

## Additional ASCII Diagrams

### System Overview
```
Users (Developers & Consultants)
              │
              ▼
      GitHub Repository
              │
   ┌──────────┼──────────┐
   ▼          ▼          ▼
macros/   scripts/    templates/
production/.bas   .ps1     ModèleImport.mpt
```

### Release Pipeline
```
/macros/production/*.bas
      │
      ▼
build_mpt.ps1
      │
      ▼
add_ribbon_to_mpt.ps1
      │
      ▼
templates/ModèleImport.mpt
```

### Developer Workflow
```
Edit bas → commit → push.ps1 → new ModèleImport.mpt
```
