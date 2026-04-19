$ErrorActionPreference = "Stop"

$projectRoot = "c:\Users\ninep\Desktop\웹디자인\배점표자동화"
$pythonExe = "C:\Users\ninep\AppData\Local\Python\pythoncore-3.14-64\python.exe"
$logDir = Join-Path $projectRoot "logs"
if (-not (Test-Path $logDir)) {
  New-Item -ItemType Directory -Path $logDir | Out-Null
}

$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$logFile = Join-Path $logDir "auto-refresh.log"

Set-Location $projectRoot

try {
  "[$stamp] start" | Out-File -FilePath $logFile -Append -Encoding utf8
  & $pythonExe "scripts/build_april_month.py" 2>&1 | Out-File -FilePath $logFile -Append -Encoding utf8
  if ($LASTEXITCODE -ne 0) {
    throw "build_april_month.py exit code: $LASTEXITCODE"
  }
  "[$stamp] done" | Out-File -FilePath $logFile -Append -Encoding utf8
} catch {
  "[$stamp] error: $($_.Exception.Message)" | Out-File -FilePath $logFile -Append -Encoding utf8
  throw
}

