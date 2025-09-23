Write-Host "== MS Project Env Check (Windows) =="

$projPaths = @(
  "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE",
  "C:\Program Files (x86)\Microsoft Office\root\Office16\WINPROJ.EXE"
)
$found = $false
foreach ($p in $projPaths) {
  if (Test-Path $p) { Write-Host "Found Project: $p"; $found = $true }
}
if (-not $found) { Write-Warning "WINPROJ.EXE not found in common locations." }

try {
  $app = New-Object -ComObject MSProject.Application
  $app.Quit()
  Write-Host "COM OK: MSProject.Application created successfully."
} catch {
  Write-Warning "COM FAILED: $($_.Exception.Message)"
}
Write-Host "Done."
