#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PidFile    = "$ProjectDir\.app.pid"

if (-not (Test-Path $PidFile)) {
    Write-Host 'Not running (no PID file)'
    exit 0
}

$AppPid = [int](Get-Content $PidFile -Raw).Trim()
$proc   = Get-Process -Id $AppPid -ErrorAction SilentlyContinue

if ($proc) {
    Stop-Process -Id $AppPid -Force
    Remove-Item $PidFile
    Write-Host "Stopped (PID: $AppPid)"
} else {
    Write-Host "Process $AppPid not found, removing stale PID file"
    Remove-Item $PidFile
}
