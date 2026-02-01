# ARCHITECTURE.md

## Overview
This document provides a **full technical explanation** of the local automation that builds and publishes a Microsoft Project template. The system is intentionally **simple, script‑driven, and CI‑free**, so developers can build and push from their workstation while consultants consume a single, signed‑off template file.

## Key Objectives
- **Consistency:** One reproducible pipeline produces predictable artifacts under `/templates/`.
- **Safety:** Macros are stored as text under `/macros/production/`, never edited inside binary Project files.
- **Simplicity:** A single command (`./scripts/push.ps1`) runs the build and publication workflow; commit/push is delegated to `commit_and_push.ps1`.

---

## Repository Structure
```text
Repository Root
 ├── macros/
 │     └── production/                # VBA source files (.bas)
 ├── scripts/                         # PowerShell automation
 │     ├── build_mpt.ps1              # Builds TemplateBase.mpt from .bas modules
 │     ├── add_ribbon_to_mpt.ps1      # Injects Ribbon; outputs TemplateBase_WithRibbon.mpt
 │     ├── commit_and_push.ps1        # Commit/push logic (dot‑sourced)
 │     └── push.ps1                   # Orchestrator
 ├── templates/
 │     ├── TemplateBase.mpt           # Output of build_mpt.ps1
 │     ├── TemplateBase_WithRibbon.mpt# Output of add_ribbon_to_mpt.ps1
 │     └── ModèleImport.mpt           # Final published artifact
 └── docs/
```

---

## Environments & Prerequisites
- **OS:** Windows (PowerShell 5+)
- **Microsoft Project:** Installed locally and licensed (required for building/testing the .mpt)
- **Git:** Installed and configured (user.name/user.email set)
- **UTF‑8 shell:** Accented filename support (`ModèleImport.mpt`)

---

## High‑Level Data Flow
```
/macros/production/*.bas
    → build_mpt.ps1
      outputs: templates/TemplateBase.mpt
    → add_ribbon_to_mpt.ps1
      TemplateBase.mpt → templates/TemplateBase_WithRibbon.mpt
    → push.ps1 → commit_and_push.ps1
      TemplateBase_WithRibbon.mpt → templates/ModèleImport.mpt (final name)
      → Git remote
```

### System Overview (ASCII)
```
Users (Developers & Consultants)
              │
              ▼
      GitHub Repository
              │
   ┌──────────┼───────────────────────┐
   ▼          ▼                       ▼
macros/   scripts/                templates/
production/.bas   .ps1    TemplateBase.mpt → TemplateBase_WithRibbon.mpt → ModèleImport.mpt
```

---

## Scripts — Detailed Behavior & Guarantees

### 🟩 build_mpt.ps1 — Template Assembly
**Purpose:** Produce a clean base template with **only** the macros defined in `/macros/production/`.

**Operations**
1. Start Microsoft Project automation context.
2. Create/open a minimal template context.
3. Purge existing modules.
4. Import modules from `/macros/production/` (deterministic order).
5. Save to:
   ```
   /templates/TemplateBase.mpt
   ```

**Guarantees:** idempotent, reproducible, isolated source of truth.

---

### 🟦 add_ribbon_to_mpt.ps1 — Ribbon Customization Injection
**Purpose:** Inject the custom Ribbon into the base template and produce the **ribbonified intermediate**.

**Inputs**
- `/templates/TemplateBase.mpt`

**Operations**
1. Open `templates/TemplateBase.mpt`.
2. Replace the **customUI** Ribbon XML.
3. Validate callbacks exist in imported modules.
4. Save the **ribbonified** template to:
   ```
   /templates/TemplateBase_WithRibbon.mpt
   ```

**Guarantees:** Full replace (no drift); callback binding safety; deterministic output.

---

### 🟧 commit_and_push.ps1 — Versioning & Remote Publish
**Purpose:** Keep commit/push logic separate for reliability and maintainability; it is dot‑sourced by `push.ps1`.

**Why Separate?**
- **PowerShell parsing edge cases** — Dot‑sourced from `push.ps1` to avoid nested try/catch parse issues when additional `try/catch` or `trap` blocks are present.
- **Clear error handling** — Uses a `trap` block for Git errors (auth, conflicts, missing remote) without interfering with the outer `try/catch` in `push.ps1`.
- **Separation of concerns** — `push.ps1` orchestrates build/ribbon/staging; commit/push logic lives here.

**Behavior**
- Stages canonical artifacts (including publishing `templates/TemplateBase_WithRibbon.mpt` as `templates/ModèleImport.mpt`)
- Commit strategy:
  - Default: **amend** previous commit
  - With `-NoAmend`: create a **new commit**
- Push to current upstream

---

### 🟥 push.ps1 — End‑to‑End Orchestrator
**Pipeline**
1. **🎨 Step 1: Injecting ribbon…** → runs `add_ribbon_to_mpt.ps1`
   - **✅ Ribbon injected successfully.** (produces `templates/TemplateBase_WithRibbon.mpt`)
2. **🔨 Step 2: Building ModèleImport.mpt…** → prepares final distributable from the ribbonified template
   - **✅ Build successful.** *Macros imported: X/X*
   - Output: `templates/ModèleImport.mpt`
3. **📦 Pushing to GitHub…** → delegates to `commit_and_push.ps1`
   - **✅ Push successful.** (amend by default; use `-NoAmend` to create a new commit)

**Console Output (indicative):**
```
🎨 Step 1: Injecting ribbon...
✅ Ribbon injected successfully.
🔨 Step 2: Building ModèleImport.mpt...
✅ Build successful. Macros imported: X/X
📦 Pushing to GitHub...
✅ Push successful.
```

**Pre‑flight Checks**
- Validate Git repo state
- Ensure no unintended untracked/unstaged changes (policy dependent)

---

## Why This Architecture Works
- **Single Source of Truth:** All code lives under `/macros/production/`.
- **Deterministic Builds:** Fixed import order + full Ribbon replacement.
- **Separation of Concerns:** Build vs. UI vs. versioning are isolated; `commit_and_push.ps1` prevents dot‑sourcing quirks.
- **Local‑First:** No CI dependency; transparent operations.

---

## Troubleshooting (Technical)
- **File lock / in use** → Close Project instances; check antivirus locks on `templates/`.
- **32‑bit vs 64‑bit Office** → Ensure correct VBA `PtrSafe` declarations.
- **Remote ahead** → `git pull --rebase`, then re‑run `commit_and_push.ps1` or `push.ps1`.
- **History too compact** → Use `-NoAmend` for explicit build commits.
