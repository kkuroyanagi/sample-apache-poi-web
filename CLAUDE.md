# CLAUDE.md

このファイルは、リポジトリ内のコードを操作する際に Claude Code (claude.ai/code) へのガイダンスを提供します。

## プロジェクト概要

Apache POI 5.3.0 を使った Excel ファイルの生成・ダウンロード機能を示す、Spring Boot 3.5.0 + React 18 の Web アプリケーション。フロントエンドはビルド時に Spring Boot JAR にバンドルされる。一部の機能は PostgreSQL（MyBatis 経由）からデータを取得する。

## ビルド・起動コマンド

```powershell
# ローカル用 PostgreSQL を起動（DB 取得系エンドポイントに必要）
docker compose up -d

# JAR をビルド（frontend-maven-plugin による React ビルドを含む）
mvn package -DskipTests

# Maven から直接起動（開発モード）
mvn spring-boot:run

# ビルドしてバックグラウンドプロセスとして起動（PowerShell）
./start.ps1

# バックグラウンドプロセスを停止
./stop.ps1

# 統合テストを実行（ポート 18080 でアプリを起動し、curl チェック後にシャットダウン）
./test_start.ps1
```

アプリへのアクセス: http://localhost:8080

## テスト

`src/test/` に **JUnit テストクラスは存在しない**。

**統合テスト（シェルスクリプト）:** ビルド・起動・HTTP レスポンス・Content-Type ヘッダー・XLSX マジックバイトを検証。
```powershell
./test_start.ps1
```

**E2E テスト（Playwright）:** `e2e/` ディレクトリで管理。アプリが http://localhost:8080 で起動している状態で実行。
```powershell
cd e2e
npx playwright test               # ヘッドレス
npx playwright test --headed      # ブラウザ表示
npx playwright test --ui          # インタラクティブ UI
```
初回のみ `npm install` と `npx playwright install chromium` が必要。

ユニットテストを追加した場合の単一テスト実行:
```
mvn test -Dtest=ClassName#methodName
```

## アーキテクチャ

**バックエンド (Java):** すべての REST コントローラーは `src/main/java/com/example/poi/controller/` に配置。各コントローラーは異なる Excel 生成シナリオを担当:

| コントローラー | エンドポイント | 用途 |
|---|---|---|
| `ExcelController` | `GET /api/excel/download` | スタイル付き単一シート売上レポート |
| `ExcelMultiSheetController` | `GET /api/excel/download-multi` | 5シート月次レポート |
| `ExcelEditableController` | `GET /api/excel/download-editable` | ドロップダウンと VLOOKUP 数式付き入力シート |
| `ExcelPerfController` | `GET /api/excel/download-xssf?rows=N` | XSSF ベンチマーク（最大 5,000 行） |
| `ExcelPerfController` | `GET /api/excel/download-sxssf?rows=N` | SXSSF ストリーミングベンチマーク（最大 100,000 行） |
| `ExcelProtectedController` | `GET /api/excel/download-protected` | シート保護デモ |
| `ExcelDbController` | `GET /api/excel/download-db` | PostgreSQL の sales テーブルから取得したデータでスタイル付き単一シートを生成 |

すべてのコントローラーは `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` と `Content-Disposition` 添付ヘッダーを持つ `ResponseEntity<byte[]>` を返す。

**フロントエンド (React/Vite):** `frontend/src/` に配置。`App.jsx` に全 UI ロジックを集約し、`useDownload()` と `usePerfDownload()` フックで blob ダウンロードを実行。Vite のビルド出力は `src/main/resources/static/` に配置され、Spring Boot から配信される。

**フロントエンドと Maven の統合:** `pom.xml` の `frontend-maven-plugin` が Node v20.18.0 を自動インストールし、`npm install` と `npm run build` を実行。Spring Boot ビルドが静的ファイルを取り込むため、フロントエンドの個別ビルド手順は不要。

**データアクセス (MyBatis + PostgreSQL):** `ExcelDbController` は `SaleMapper`（`src/main/java/com/example/poi/mapper/`）経由で `sales` テーブルを読み、`Sale` レコード（`domain/`）にマッピングする。MyBatis は `map-underscore-to-camel-case` により snake_case カラムを camelCase フィールドへ変換。テーブル定義と初期データは `src/main/resources/schema.sql` / `data.sql`。ローカルの PostgreSQL は `compose.yml`（`docker compose up -d`）で起動。

## 技術的な重要事項

- **XSSF vs SXSSF:** `ExcelPerfController` でインメモリ (XSSF) とストリーミング (SXSSF) の Excel 生成をベンチマーク。SXSSF は行をテンポラリファイルに書き出してメモリ上にスライディングウィンドウのみ保持するため、大規模データセットに適している。パフォーマンス指標はレスポンスヘッダー (`X-Row-Count`、`X-Generation-Time-Ms`、`X-Heap-Delta-MB`) で返される。
- **非同期タイムアウト:** `application.yml` の `spring.mvc.async.request-timeout: 300000`（5分）により、大きな Excel ファイルの生成に対応。
- **ポート:** デフォルト 8080。スクリプトは `PORT=XXXX` での上書きをサポート。
- **プロファイルと DB 接続切替:** `SPRING_PROFILES_ACTIVE` 環境変数でプロファイルを選択（未指定なら `local`）。`application-local.yml` は `compose.yml` の PostgreSQL を既定値で参照し、起動時に `schema.sql`/`data.sql` を投入（`spring.sql.init.mode: always`）。`application-test.yml` は認証情報を環境変数 `DB_URL` / `DB_USERNAME` / `DB_PASSWORD`（既定値なし＝必須）から受け取り、DB 初期化は行わない（`mode: never`）。テストサーバーでは `SPRING_PROFILES_ACTIVE=test` と各 `DB_*` を設定して起動する。各プロファイルの接続情報も同名の環境変数で個別に上書き可能。
