#Requires -Version 5.1
$ErrorActionPreference = 'SilentlyContinue'

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Port       = 18080   # テスト用ポート（本番 8080 と分離）
$BaseUrl    = "http://localhost:$Port"
$Pass       = 0
$Fail       = 0
$PidFile    = "$ProjectDir\.app.pid"

function Write-Pass([string]$Label) { Write-Host "  [PASS] $Label"; $script:Pass++ }
function Write-Fail([string]$Label) { Write-Host "  [FAIL] $Label"; $script:Fail++ }

function Assert-Equal([string]$Label, $Expected, $Actual) {
    if ($Expected -eq $Actual) { Write-Pass $Label }
    else { Write-Fail "$Label (expected: $Expected, actual: $Actual)" }
}

function Assert-Contains([string]$Label, [string]$Needle, [string]$Haystack) {
    if ($Haystack -match [regex]::Escape($Needle)) { Write-Pass $Label }
    else { Write-Fail "$Label (expected to contain: $Needle)" }
}

try {
    Write-Host "=== 起動スクリプト 統合テスト ==="
    Write-Host ""

    # --- T1: start.ps1 が JAR をビルドしてプロセスを起動する ---
    Write-Host "T1: start.ps1 — ビルド & 起動"
    $env:PORT = $Port
    & "$ProjectDir\start.ps1"

    $appId = if (Test-Path $PidFile) { [int](Get-Content $PidFile -Raw).Trim() } else { 0 }
    if ($appId -and (Get-Process -Id $appId -ErrorAction SilentlyContinue)) {
        Write-Pass "PID ファイルが作成され、プロセスが存在する"
    } else {
        Write-Fail "PID ファイルが存在しないか、プロセスが起動していない"
    }

    # --- T2: アプリが起動するまで待機（最大 60 秒）---
    Write-Host "T2: アプリ起動待機（最大 60 秒）"
    $ready = $false
    foreach ($i in 1..60) {
        try {
            $r = Invoke-WebRequest -Uri "$BaseUrl/api/excel/download" -UseBasicParsing -TimeoutSec 2
            if ($r.StatusCode -eq 200) { $ready = $true; break }
        } catch {}
        Start-Sleep -Seconds 1
    }

    if ($ready) {
        Write-Pass "60 秒以内に起動完了"
    } else {
        Write-Fail "タイムアウト: アプリが起動しなかった"
        Write-Host "ログ:"
        Get-Content "$ProjectDir\.app.log" -Tail 20 -ErrorAction SilentlyContinue
        exit 1
    }

    # --- T3: 二重起動が抑止される ---
    Write-Host "T3: 二重起動の抑止"
    $env:PORT = $Port
    $output = & "$ProjectDir\start.ps1" 2>&1
    Assert-Contains "Already running メッセージが表示される" "Already running" ($output -join " ")

    # --- T4: /api/excel/download が 200 を返す ---
    Write-Host "T4: Excel ダウンロード API — HTTP 200"
    try {
        $r = Invoke-WebRequest -Uri "$BaseUrl/api/excel/download" -UseBasicParsing
        Assert-Equal "HTTP ステータスが 200" 200 $r.StatusCode
    } catch { Write-Fail "HTTP ステータスが 200 (例外: $_)" }

    # --- T5: Content-Type が xlsx ---
    Write-Host "T5: Excel ダウンロード API — Content-Type"
    try {
        $r  = Invoke-WebRequest -Uri "$BaseUrl/api/excel/download" -UseBasicParsing
        $ct = $r.Headers.'Content-Type'
        Assert-Contains "Content-Type に spreadsheetml が含まれる" "spreadsheetml" $ct
    } catch { Write-Fail "Content-Type の確認失敗 (例外: $_)" }

    # --- T6: Content-Disposition に sales_report.xlsx が含まれる ---
    Write-Host "T6: Excel ダウンロード API — Content-Disposition"
    try {
        $r  = Invoke-WebRequest -Uri "$BaseUrl/api/excel/download" -UseBasicParsing
        $cd = $r.Headers.'Content-Disposition'
        Assert-Contains "Content-Disposition に sales_report.xlsx が含まれる" "sales_report.xlsx" $cd
    } catch { Write-Fail "Content-Disposition の確認失敗 (例外: $_)" }

    # --- T7: ダウンロードしたファイルが xlsx マジックバイト（PK）で始まる ---
    Write-Host "T7: ダウンロードファイルが有効な xlsx（ZIP 形式）"
    try {
        $tmp   = [System.IO.Path]::GetTempFileName()
        Invoke-WebRequest -Uri "$BaseUrl/api/excel/download" -OutFile $tmp -UseBasicParsing
        $bytes = [System.IO.File]::ReadAllBytes($tmp)[0..1]
        $magic = ($bytes | ForEach-Object { $_.ToString('x2') }) -join ''
        Assert-Equal "ファイルが PK（ZIP マジックバイト）で始まる" "504b" $magic
        Remove-Item $tmp -ErrorAction SilentlyContinue
    } catch { Write-Fail "ファイル検証失敗 (例外: $_)" }

    # --- T8: stop.ps1 がプロセスを停止する ---
    Write-Host "T8: stop.ps1 — プロセス停止"
    & "$ProjectDir\stop.ps1"
    Start-Sleep -Seconds 2
    if (-not (Test-Path $PidFile)) {
        Write-Pass "PID ファイルが削除された"
    } else {
        Write-Fail "PID ファイルが残っている"
    }

    # --- T9: 停止後にポートが解放される ---
    Write-Host "T9: 停止後のポート解放"
    try {
        Invoke-WebRequest -Uri "$BaseUrl/api/excel/download" -UseBasicParsing -TimeoutSec 2 | Out-Null
        Write-Fail "ポート $Port がまだ応答している"
    } catch {
        Write-Pass "ポート $Port が解放された"
    }

} finally {
    try { & "$ProjectDir\stop.ps1" } catch {}
}

Write-Host ""
Write-Host "=== 結果: PASS=$Pass  FAIL=$Fail ==="
if ($Fail -eq 0) { exit 0 } else { exit 1 }
