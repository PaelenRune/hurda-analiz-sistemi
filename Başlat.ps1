$ErrorActionPreference = 'Stop'
$taskRoot = $PSScriptRoot
$taskPython = Join-Path $taskRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $taskPython)) {
    throw 'Python ortamı bulunamadı. KURULUM.md dosyasındaki kurulum adımlarını uygulayın.'
}
$taskPort = 8770
try {
    $taskResponse = Invoke-WebRequest -Uri "http://127.0.0.1:$taskPort/api/state" -TimeoutSec 2 -ErrorAction Stop
} catch {
    Start-Process -FilePath $taskPython -ArgumentList @('app.py','--port',"$taskPort") -WorkingDirectory $taskRoot -WindowStyle Hidden
    for ($taskAttempt=0;$taskAttempt -lt 40;$taskAttempt++) {
        Start-Sleep -Milliseconds 500
        try {$taskResponse=Invoke-WebRequest -Uri "http://127.0.0.1:$taskPort/api/state" -TimeoutSec 1 -ErrorAction Stop;break} catch {}
    }
}
Start-Process "http://127.0.0.1:$taskPort/"
