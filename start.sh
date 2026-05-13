#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
JAR="$PROJECT_DIR/target/sample-apache-poi-web-0.0.1-SNAPSHOT.jar"
PID_FILE="$PROJECT_DIR/.app.pid"
PORT="${PORT:-8080}"

# Java 21+ を探す（pom.xml の要件）
find_java() {
  # 環境変数 JAVA_HOME が設定済みで 21+ なら使う
  if [ -n "${JAVA_HOME:-}" ] && "$JAVA_HOME/bin/java" -version 2>&1 | grep -qE '"(2[1-9]|[3-9][0-9]|[0-9]{3,})'; then
    echo "$JAVA_HOME/bin/java"; return
  fi
  # Homebrew openjdk（Apple Silicon / Intel）
  for candidate in \
      /opt/homebrew/opt/openjdk/bin/java \
      /usr/local/opt/openjdk/bin/java; do
    if [ -x "$candidate" ] && "$candidate" -version 2>&1 | grep -qE '"(2[1-9]|[3-9][0-9]|[0-9]{3,}|21)'; then
      echo "$candidate"; return
    fi
  done
  # PATH 上の java が 21+ か確認
  if command -v java &>/dev/null && java -version 2>&1 | grep -qE '"(2[1-9]|[3-9][0-9]|[0-9]{3,}|21)'; then
    echo "java"; return
  fi
  echo ""
}

JAVA_CMD=$(find_java)
if [ -z "$JAVA_CMD" ]; then
  echo "Error: Java 21 以上が見つかりません。JAVA_HOME を設定するか、Java 21+ をインストールしてください。" >&2
  exit 1
fi

# すでに起動中か確認
if [ -f "$PID_FILE" ] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
  echo "Already running (PID: $(cat "$PID_FILE")) — http://localhost:$PORT"
  exit 0
fi

# JAR がなければビルド
if [ ! -f "$JAR" ]; then
  echo "Building..."
  mvn -f "$PROJECT_DIR/pom.xml" package -DskipTests -q
fi

echo "Starting with: $JAVA_CMD"
"$JAVA_CMD" -jar "$JAR" --server.port="$PORT" \
  > "$PROJECT_DIR/.app.log" 2>&1 &
echo $! > "$PID_FILE"
echo "Started (PID: $(cat "$PID_FILE")) — http://localhost:$PORT"
echo "Log: $PROJECT_DIR/.app.log"
