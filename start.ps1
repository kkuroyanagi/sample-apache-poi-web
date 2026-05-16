#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Jar        = "$ProjectDir\target\sample-apache-poi-web-0.0.1-SNAPSHOT.jar"
$PidFile    = "$ProjectDir\.app.pid"
$Port       = if ($env:PORT) { $env:PORT } else { '8080' }
$LogFile    = "$ProjectDir\.app.log"

function Get-JavaMajorVersion([string]$JavaExe) {
    # java -version writes to stderr; NativeCommandError objects cause terminating errors
    # when $ErrorActionPreference = 'Stop', so override it locally.
    $local:ErrorActionPreference = 'Continue'
    try {
        $lines = & $JavaExe -version 2>&1 | ForEach-Object { "$_" }
        $line  = $lines | Where-Object { $_ -match 'version "' } | Select-Object -First 1
        if ($line -match 'version "(\d+)') { return [int]$Matches[1] }
    } catch {}
    return 0
}

function Find-Java {
    if ($env:JAVA_HOME) {
        $c = Join-Path $env:JAVA_HOME 'bin\java.exe'
        if ((Test-Path $c) -and (Get-JavaMajorVersion $c) -ge 21) { return $c }
    }
    if (Get-Command java -ErrorAction SilentlyContinue) {
        if ((Get-JavaMajorVersion 'java') -ge 21) { return 'java' }
    }
    return $null
}

$JavaCmd = Find-Java
if (-not $JavaCmd) {
    Write-Error 'Error: Java 21 以上が見つかりません。JAVA_HOME を設定するか、Java 21+ をインストールしてください。'
    exit 1
}

if (Test-Path $PidFile) {
    $existingId = (Get-Content $PidFile -Raw).Trim()
    if ($existingId -and (Get-Process -Id ([int]$existingId) -ErrorAction SilentlyContinue)) {
        Write-Host "Already running (PID: $existingId) — http://localhost:$Port"
        exit 0
    }
}

if (-not (Test-Path $Jar)) {
    Write-Host 'Building...'
    & mvn -f "$ProjectDir\pom.xml" package -DskipTests -q
    if ($LASTEXITCODE -ne 0) { Write-Error 'ビルドに失敗しました。'; exit 1 }
}

Write-Host "Starting with: $JavaCmd"
$proc = Start-Process -FilePath $JavaCmd `
    -ArgumentList "-jar `"$Jar`" --server.port=$Port" `
    -RedirectStandardOutput $LogFile `
    -RedirectStandardError  "$LogFile.err" `
    -NoNewWindow -PassThru

$proc.Id | Out-File -FilePath $PidFile -Encoding ascii -NoNewline
Write-Host "Started (PID: $($proc.Id)) — http://localhost:$Port"
Write-Host "Log:    $LogFile"
Write-Host "ErrLog: $LogFile.err"