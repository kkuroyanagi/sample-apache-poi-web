package com.example.poi.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.management.ManagementFactory;

@RestController
@RequestMapping("/api/excel")
public class ExcelPerfController {

    static final int TOTAL_COLUMNS   = 300;
    static final int INPUT_COLS      = 10;    // 先頭 10 列がプルダウン入力列
    static final int XSSF_ROW_LIMIT  = 5_000; // 300列×5K = 1.5M セル
    static final int MAX_ROWS        = 100_000;

    /**
     * マスタデータ: index 0 = 列ヘッダー名、index 1以降 = 選択肢
     * INPUT_COLS（10）個のリストを定義。
     */
    static final String[][] MASTER_DATA = {
        {"カテゴリ",      "電子機器", "周辺機器", "記憶媒体", "ネットワーク機器", "ソフトウェア"},
        {"ステータス",    "受注", "出荷済", "請求済", "完了", "キャンセル"},
        {"地域",          "北海道", "東北", "関東", "中部", "近畿", "中国", "四国", "九州"},
        {"担当者",        "田中", "鈴木", "佐藤", "高橋", "伊藤", "渡辺"},
        {"優先度",        "高", "中", "低"},
        {"部門",          "営業部", "技術部", "管理部", "マーケティング部", "財務部"},
        {"支払方法",      "現金", "振込", "クレジットカード", "電子マネー"},
        {"配送方法",      "標準配送", "特急配送", "翌日配送", "店頭受取"},
        {"承認状態",      "未承認", "承認済", "保留", "却下"},
        {"品質チェック",  "合格", "不合格", "検査中", "免除"},
    };

    // ── XSSF（全行をメモリに保持）────────────────────────────────────────────

    @GetMapping("/download-xssf")
    public ResponseEntity<byte[]> downloadXssf(
            @RequestParam(defaultValue = "1000") int rows) throws IOException {

        rows = Math.min(rows, MAX_ROWS);
        if (rows > XSSF_ROW_LIMIT) {
            String msg = "{\"error\":\"XSSF は %,d 行以下で指定してください（300列モード）。それ以上は SXSSF を使用してください。\"}"
                .formatted(XSSF_ROW_LIMIT);
            return ResponseEntity.badRequest()
                .contentType(MediaType.APPLICATION_JSON)
                .body(msg.getBytes());
        }

        System.gc();
        long heapBefore = heapUsed();
        long startMs    = System.currentTimeMillis();

        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            buildWorkbook(wb, rows);
            wb.write(out);
            bytes = out.toByteArray();
        } catch (OutOfMemoryError e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .contentType(MediaType.APPLICATION_JSON)
                .body("{\"error\":\"OutOfMemoryError: 行数を減らすか SXSSF を使用してください。\"}".getBytes());
        }

        long genMs     = System.currentTimeMillis() - startMs;
        long heapDelta = heapUsed() - heapBefore;

        return ResponseEntity.ok()
            .headers(buildHeaders("perf_xssf.xlsx", rows, genMs, heapDelta))
            .body(bytes);
    }

    // ── SXSSF（ディスクストリーミング・省メモリ）─────────────────────────────

    @GetMapping("/download-sxssf")
    public ResponseEntity<byte[]> downloadSxssf(
            @RequestParam(defaultValue = "1000") int rows) throws IOException {

        rows = Math.min(rows, MAX_ROWS);

        System.gc();
        long heapBefore = heapUsed();
        long startMs    = System.currentTimeMillis();

        final long genMs;
        final long heapDelta;
        byte[] bytes;

        try (SXSSFWorkbook wb = new SXSSFWorkbook(1000);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            buildWorkbook(wb, rows);
            genMs     = System.currentTimeMillis() - startMs;
            heapDelta = heapUsed() - heapBefore;
            wb.write(out);
            bytes = out.toByteArray();
        }

        return ResponseEntity.ok()
            .headers(buildHeaders("perf_sxssf.xlsx", rows, genMs, heapDelta))
            .body(bytes);
    }

    // ── ワークブック構築（XSSF/SXSSF 共通）──────────────────────────────────

    private void buildWorkbook(Workbook wb, int rows) {
        // 入力シートを先に作成してシートインデックス 0 にする
        Sheet dataSheet   = wb.createSheet("パフォーマンステスト");
        Sheet masterSheet = wb.createSheet("マスタ");
        wb.setActiveSheet(0);

        writeMasterSheet(wb, masterSheet);
        writeDataSheet(wb, dataSheet, rows);
    }

    // ── Sheet: マスタ ─────────────────────────────────────────────────────────

    private void writeMasterSheet(Workbook wb, Sheet sheet) {
        CellStyle headerStyle = buildHeaderStyle(wb);
        CellStyle valStyle    = buildStrStyle(wb);

        // 列幅
        for (int col = 0; col < INPUT_COLS; col++) sheet.setColumnWidth(col, 5000);

        // ヘッダー行（MASTER_DATA[i][0]）
        Row hRow = sheet.createRow(0);
        for (int col = 0; col < INPUT_COLS; col++) {
            Cell c = hRow.createCell(col);
            c.setCellValue(MASTER_DATA[col][0]);
            c.setCellStyle(headerStyle);
        }

        // 選択肢行（MASTER_DATA[i][1+]）
        int maxOptions = 0;
        for (String[] list : MASTER_DATA) maxOptions = Math.max(maxOptions, list.length - 1);

        for (int row = 1; row <= maxOptions; row++) {
            Row r = sheet.createRow(row);
            for (int col = 0; col < INPUT_COLS; col++) {
                String[] list = MASTER_DATA[col];
                if (row < list.length) {
                    Cell c = r.createCell(col);
                    c.setCellValue(list[row]);
                    c.setCellStyle(valStyle);
                }
            }
        }
    }

    // ── Sheet: パフォーマンステスト ───────────────────────────────────────────

    private void writeDataSheet(Workbook wb, Sheet sheet, int rows) {
        sheet.setDefaultColumnWidth(7);
        for (int col = 0; col < INPUT_COLS; col++) sheet.setColumnWidth(col, 4500);

        // スタイルをループ外で一度だけ生成
        CellStyle headerStyle = buildHeaderStyle(wb);
        CellStyle inputStyle  = buildInputStyle(wb);
        CellStyle numStyle    = buildNumStyle(wb);

        // ── ヘッダー行 ──
        Row hRow = sheet.createRow(0);
        for (int col = 0; col < TOTAL_COLUMNS; col++) {
            Cell c = hRow.createCell(col);
            c.setCellValue(col < INPUT_COLS
                ? MASTER_DATA[col][0]
                : "D_%03d".formatted(col + 1));
            c.setCellStyle(headerStyle);
        }

        // ── データ行 ──
        for (int i = 0; i < rows; i++) {
            Row r = sheet.createRow(i + 1);

            // 入力列（0〜9）: マスタの選択肢を循環セット
            for (int col = 0; col < INPUT_COLS; col++) {
                String[] opts  = MASTER_DATA[col];
                int optCount   = opts.length - 1;           // 選択肢数（ヘッダー除く）
                Cell c = r.createCell(col);
                c.setCellValue(opts[1 + (i % optCount)]);   // opts[1]〜opts[n] を循環
                c.setCellStyle(inputStyle);
            }

            // データ列（10〜299）: 連番数値
            for (int col = INPUT_COLS; col < TOTAL_COLUMNS; col++) {
                Cell c = r.createCell(col);
                c.setCellValue((long) i * TOTAL_COLUMNS + col);
                c.setCellStyle(numStyle);
            }
        }

        // ── プルダウン（DataValidation）──
        addDataValidations(sheet, rows);
    }

    /**
     * 入力列 10 本にそれぞれ DataValidation を設定。
     * リスト参照先: 'マスタ'!$A$2:$A$N (各列の選択肢数に合わせた範囲)
     */
    private void addDataValidations(Sheet sheet, int rows) {
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();

        for (int col = 0; col < INPUT_COLS; col++) {
            String[] list    = MASTER_DATA[col];
            int optCount     = list.length - 1;                     // 選択肢数
            String colLetter = toColumnLetter(col);                  // A〜J
            // 選択肢が入る Excel 行範囲: row 2 〜 row (optCount+1)
            String formula = "'マスタ'!$%s$2:$%s$%d"
                .formatted(colLetter, colLetter, optCount + 1);

            DataValidationConstraint constraint =
                dvHelper.createFormulaListConstraint(formula);
            // POI 行インデックス 1〜rows（ヘッダー除くデータ行全体）
            CellRangeAddressList range = new CellRangeAddressList(1, rows, col, col);

            DataValidation dv = dvHelper.createValidation(constraint, range);
            dv.setShowErrorBox(true);
            dv.setErrorStyle(DataValidation.ErrorStyle.WARNING);
            dv.createErrorBox("入力エラー", "マスタにない値です。プルダウンから選択してください。");
            sheet.addValidationData(dv);
        }
    }

    /** 0 始まり列インデックス → Excel 列文字列（0=A, 1=B, …, 25=Z, 26=AA）*/
    private static String toColumnLetter(int col) {
        StringBuilder sb = new StringBuilder();
        for (int c = col + 1; c > 0; c = (c - 1) / 26)
            sb.insert(0, (char) ('A' + (c - 1) % 26));
        return sb.toString();
    }

    // ── ヘルパー ─────────────────────────────────────────────────────────────

    private static long heapUsed() {
        return ManagementFactory.getMemoryMXBean().getHeapMemoryUsage().getUsed();
    }

    private HttpHeaders buildHeaders(String filename, int rows, long genMs, long heapDeltaBytes) {
        HttpHeaders h = new HttpHeaders();
        h.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        h.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"");
        h.set("X-Row-Count",          String.valueOf(rows));
        h.set("X-Generation-Time-Ms", String.valueOf(genMs));
        h.set("X-Heap-Delta-MB",      String.valueOf(heapDeltaBytes / 1024 / 1024));
        return h;
    }

    // ── スタイル ─────────────────────────────────────────────────────────────

    private CellStyle buildHeaderStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true); f.setColor(IndexedColors.WHITE.getIndex()); s.setFont(f);
        s.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        return s;
    }

    private CellStyle buildInputStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.LEFT);
        return s;
    }

    private CellStyle buildStrStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setAlignment(HorizontalAlignment.LEFT);
        return s;
    }

    private CellStyle buildNumStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0"));
        s.setAlignment(HorizontalAlignment.RIGHT);
        return s;
    }
}
