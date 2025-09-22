# GitHub Downloader (VBA) — Usage & Troubleshooting

## What this module does
- Downloads files over HTTPS from GitHub (Raw host or GitHub Pages).
- Logs progress to the Immediate Window (`Ctrl+G`) with timestamps.
- Works on Windows and Ubuntu/Wine:
  - Saves to `~/Downloads/omexom` (Linux) or `%USERPROFILE%\Downloads\omexom` (Windows).
  - Falls back to `curl` if COM HTTP is blocked.

## Files
- `excel/ModuleGitHubDownload.bas`  ← exported VBA module

## Quick start (guaranteed success test)
We first download the repo README (always present on `main`) to prove the pipeline works.

1. Open **VBE** (Alt+F11) → open module **ModuleGitHubDownload**.
2. Ensure these two lines are set as below:
   ```vba
   Private Const BASE_URL As String = "https://raw.githubusercontent.com/lfa-lab/Omexom/main/"
   ' ...
   relative = "README.md"
