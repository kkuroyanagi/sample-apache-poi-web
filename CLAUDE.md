# CLAUDE.md

このファイルは、リポジトリ内のコードを操作する際に Claude Code (claude.ai/code) へのガイダンスを提供します。

## プロジェクト概要

Apache POI 5.3.0 を使った Excel ファイルの生成・ダウンロード機能を示す、Spring Boot 3.5.0 + React 18 の Web アプリケーション。フロントエンドはビルド時に Spring Boot JAR にバンドルされる。

## ビルド・起動コマンド

```powershell
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

`src/test/` に **JUnit テストクラスは存在しない**。テストはすべて統合スクリプトで実施:

```powershell
./test_start.ps1   # ビルド・起動・HTTPレスポンス・Content-Typeヘッダー・XLSXマジックバイトを検証
```

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

すべてのコントローラーは `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` と `Content-Disposition` 添付ヘッダーを持つ `ResponseEntity<byte[]>` を返す。

**フロントエンド (React/Vite):** `frontend/src/` に配置。`App.jsx` に全 UI ロジックを集約し、`useDownload()` と `usePerfDownload()` フックで blob ダウンロードを実行。Vite のビルド出力は `src/main/resources/static/` に配置され、Spring Boot から配信される。

**フロントエンドと Maven の統合:** `pom.xml` の `frontend-maven-plugin` が Node v20.18.0 を自動インストールし、`npm install` と `npm run build` を実行。Spring Boot ビルドが静的ファイルを取り込むため、フロントエンドの個別ビルド手順は不要。

## 技術的な重要事項

- **XSSF vs SXSSF:** `ExcelPerfController` でインメモリ (XSSF) とストリーミング (SXSSF) の Excel 生成をベンチマーク。SXSSF は行をテンポラリファイルに書き出してメモリ上にスライディングウィンドウのみ保持するため、大規模データセットに適している。パフォーマンス指標はレスポンスヘッダー (`X-Row-Count`、`X-Generation-Time-Ms`、`X-Heap-Delta-MB`) で返される。
- **非同期タイムアウト:** `application.properties` の `spring.mvc.async.request-timeout=300000`（5分）により、大きな Excel ファイルの生成に対応。
- **ポート:** デフォルト 8080。スクリプトは `PORT=XXXX` での上書きをサポート。
