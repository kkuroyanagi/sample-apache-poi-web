# sample-apache-poi-web

Apache POI を使って Excel ファイルを生成し、Web ブラウザからダウンロードできるサンプルアプリケーションです。

## 構成

| レイヤー | 技術 |
|---|---|
| バックエンド | Spring Boot 3.5.0 / Java 21 |
| Excel 生成 | Apache POI 5.3.0 (poi-ooxml) |
| フロントエンド | React 18 / Vite 5 |
| ビルド | Maven 3.6+ + frontend-maven-plugin |

## 機能

| 機能 | 説明 |
|---|---|
| 単票レポート | 売上データ（2025年4月）を 1 シートで出力 |
| 5 シートレポート | 月別明細・カテゴリ集計・月次比較・サマリーを複数シートで出力 |
| 入力シート | マスタ参照プルダウン付き編集シート（VLOOKUP 自動入力） |
| 保護シート | カテゴリ・商品マスタシートを保護した入力シート |
| パフォーマンステスト | XSSF（全行ヒープ保持）vs SXSSF（ストリーミング）を 300 列で比較計測 |

## 動作要件

- Java 21 以上
- Maven 3.6 以上
- Node.js は frontend-maven-plugin が自動インストール（v20.18.0）するため、アプリのビルド・起動には不要
- Playwright e2e テストを実行する場合は Node.js 18 以上が必要

## 起動方法

### スクリプトを使う（推奨）

```bash
# ビルド + バックグラウンド起動
./start.sh        # macOS / Linux
./start.ps1       # Windows (PowerShell)

# 停止
./stop.sh
./stop.ps1
```

`start.sh` / `start.ps1` は Java 21+ を自動検出し、JAR が存在しない場合はビルドも実行します。  
ログは `.app.log`、PID は `.app.pid` に保存されます。

起動後、ブラウザで http://localhost:8080 を開いてください。

### Maven で直接起動する

```bash
mvn spring-boot:run
```

### ポートを変更する

```bash
PORT=9090 ./start.sh
```

## ビルド

```bash
mvn package -DskipTests
```

`target/sample-apache-poi-web-0.0.1-SNAPSHOT.jar` が生成されます。

## テスト

### 統合テスト（シェルスクリプト）

```bash
./test_start.sh    # macOS / Linux
./test_start.ps1   # Windows (PowerShell)
```

9 ケースを検証します（起動・HTTP ステータス・Content-Type・xlsx マジックバイト・停止など）。

### E2E テスト（Playwright）

```bash
cd e2e
npm install
npx playwright install chromium   # 初回のみ
npx playwright test               # ヘッドレス実行
npx playwright test --headed      # ブラウザを表示して実行
npx playwright test --ui          # インタラクティブ UI モード
```

アプリが http://localhost:8080 で起動している状態で実行してください。

## API

| メソッド | パス | 説明 |
|---|---|---|
| GET | `/api/excel/download` | 単票売上レポート（1 シート）|
| GET | `/api/excel/download-multi` | 5 シートレポート（月別・集計・比較・サマリー）|
| GET | `/api/excel/download-editable` | マスタ選択プルダウン付き入力シート |
| GET | `/api/excel/download-protected` | マスタシート保護付き入力シート |
| GET | `/api/excel/download-xssf?rows=N` | XSSF パフォーマンステスト（上限 5,000 行・300 列）|
| GET | `/api/excel/download-sxssf?rows=N` | SXSSF パフォーマンステスト（上限 100,000 行・300 列）|

### パフォーマンステスト パラメータ

| パラメータ | デフォルト | 説明 |
|---|---|---|
| `rows` | 1000 | 生成行数 |

レスポンスヘッダーで計測値を返します。

| ヘッダー | 説明 |
|---|---|
| `X-Row-Count` | 実際の生成行数 |
| `X-Generation-Time-Ms` | サーバー側の Excel 生成時間（ms）|
| `X-Heap-Delta-MB` | 生成前後のヒープ増減（MB）|

## Apache POI 使用機能

| 機能 | 使用箇所 |
|---|---|
| セルスタイル（背景色・フォント・罫線・数値書式）| 全エンドポイント |
| セル結合（`CellRangeAddress`）| 全エンドポイント |
| 複数シート（`createSheet`）| 5 シート・入力シート・パフォーマンステスト |
| データ入力規則（`DataValidation`）| 入力シート・パフォーマンステスト |
| VLOOKUP 数式（`setCellFormula`）| 入力シート |
| シート保護（`sheet.protectSheet`）| 保護シート |
| ストリーミング書き込み（`SXSSFWorkbook`）| パフォーマンステスト |

## プロジェクト構成

```
.
├── start.sh / start.ps1            # 起動スクリプト
├── stop.sh / stop.ps1              # 停止スクリプト
├── test_start.sh / test_start.ps1  # 統合テスト
├── pom.xml
├── src/main/
│   ├── java/com/example/poi/
│   │   ├── PoiWebApplication.java
│   │   └── controller/
│   │       ├── ExcelController.java             # 単票レポート
│   │       ├── ExcelMultiSheetController.java   # 5 シートレポート
│   │       ├── ExcelEditableController.java     # 入力シート（プルダウン）
│   │       ├── ExcelProtectedController.java    # 保護シート
│   │       └── ExcelPerfController.java         # XSSF/SXSSF 比較
│   └── resources/
│       └── application.properties
├── frontend/
│   ├── src/
│   │   ├── App.jsx                 # メイン UI
│   │   └── App.css
│   └── package.json
└── e2e/                            # Playwright e2e テスト
    ├── tests/
    │   └── basic.spec.js
    ├── playwright.config.js
    └── package.json
```
