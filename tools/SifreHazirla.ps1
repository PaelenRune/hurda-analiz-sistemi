$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$taskPython = Join-Path $taskRoot '.venv\Scripts\python.exe'
& $taskPython (Join-Path $PSScriptRoot 'password_hash.py')
if ($LASTEXITCODE -ne 0) { Write-Host 'İşlem tamamlanamadı. Pencereyi kapatıp tekrar çalıştırın.' }
