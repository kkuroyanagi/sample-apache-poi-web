#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
PORT=18080  # テスト用ポート（本番 8080 と分離）
BASE_URL="http://localhost:$PORT"
PASS=0
FAIL=0

# ── ヘルパー ──────────────────────────────────────────────────────────────────

pass() { echo "  [PASS] $1"; PASS=$((PASS + 1)); }
fail() { echo "  [FAIL] $1"; FAIL=$((FAIL + 1)); }

assert_eq() {
  local label="$1" expected="$2" actual="$3"
  if [ "$expected" = "$actual" ]; then pass "$label"; else fail "$label (expected: $expected, actual: $actual)"; fi
}

assert_contains() {
  local label="$1" needle="$2" haystack="$3"
  if echo "$haystack" | grep -q "$needle"; then pass "$label"; else fail "$label (expected to contain: $needle)"; fi
}

cleanup() {
  PORT=$PORT "$PROJECT_DIR/stop.sh" 2>/dev/null || true
}
trap cleanup EXIT

# ── テスト ────────────────────────────────────────────────────────────────────

echo "=== 起動スクリプト 統合テスト ==="
echo

# --- T1: start.sh が JAR をビルドしてプロセスを起動する ---
echo "T1: start.sh — ビルド & 起動"
PORT=$PORT "$PROJECT_DIR/start.sh"

PID_FILE="$PROJECT_DIR/.app.pid"
if [ -f "$PID_FILE" ] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
  pass "PID ファイルが作成され、プロセスが存在する"
else
  fail "PID ファイルが存在しないか、プロセスが起動していない"
fi

# --- T2: アプリが起動するまで待機（最大 60 秒）---
echo "T2: アプリ起動待機（最大 60 秒）"
READY=false
for i in $(seq 1 60); do
  if curl -sf "$BASE_URL/api/excel/download" -o /dev/null -w "%{http_code}" 2>/dev/null | grep -q "200"; then
    READY=true
    break
  fi
  sleep 1
done

if $READY; then
  pass "60 秒以内に起動完了"
else
  fail "タイムアウト: アプリが起動しなかった"
  echo "ログ:"
  tail -20 "$PROJECT_DIR/.app.log" 2>/dev/null || true
  exit 1
fi

# --- T3: 二重起動が抑止される ---
echo "T3: 二重起動の抑止"
OUTPUT=$(PORT=$PORT "$PROJECT_DIR/start.sh" 2>&1)
assert_contains "Already running メッセージが表示される" "Already running" "$OUTPUT"

# --- T4: /api/excel/download が 200 を返す ---
echo "T4: Excel ダウンロード API — HTTP 200"
STATUS=$(curl -sf -o /dev/null -w "%{http_code}" "$BASE_URL/api/excel/download")
assert_eq "HTTP ステータスが 200" "200" "$STATUS"

# --- T5: Content-Type が xlsx ---
echo "T5: Excel ダウンロード API — Content-Type"
CT=$(curl -sf -I "$BASE_URL/api/excel/download" | grep -i "content-type:" | tr -d '\r')
assert_contains "Content-Type に spreadsheetml が含まれる" "spreadsheetml" "$CT"

# --- T6: Content-Disposition に sales_report.xlsx が含まれる ---
echo "T6: Excel ダウンロード API — Content-Disposition"
CD=$(curl -sf -I "$BASE_URL/api/excel/download" | grep -i "content-disposition:" | tr -d '\r')
assert_contains "Content-Disposition に sales_report.xlsx が含まれる" "sales_report.xlsx" "$CD"

# --- T7: ダウンロードしたファイルが xlsx マジックバイト（PK）で始まる ---
echo "T7: ダウンロードファイルが有効な xlsx（ZIP 形式）"
TMPFILE=$(mktemp /tmp/test_xlsx_XXXXXX)
curl -sf "$BASE_URL/api/excel/download" -o "$TMPFILE"
MAGIC=$(xxd -p -l 2 "$TMPFILE")
assert_eq "ファイルが PK（ZIP マジックバイト）で始まる" "504b" "$MAGIC"
rm -f "$TMPFILE"

# --- T8: stop.sh がプロセスを停止する ---
echo "T8: stop.sh — プロセス停止"
PORT=$PORT "$PROJECT_DIR/stop.sh"
sleep 2
if [ ! -f "$PID_FILE" ]; then
  pass "PID ファイルが削除された"
else
  fail "PID ファイルが残っている"
fi

# --- T9: 停止後にポートが解放される ---
echo "T9: 停止後のポート解放"
if curl -sf --max-time 2 "$BASE_URL/api/excel/download" -o /dev/null 2>/dev/null; then
  fail "ポート $PORT がまだ応答している"
else
  pass "ポート $PORT が解放された"
fi

# ── 結果 ─────────────────────────────────────────────────────────────────────

echo
echo "=== 結果: PASS=$PASS  FAIL=$FAIL ==="
[ "$FAIL" -eq 0 ] && exit 0 || exit 1
