package com.example.poi.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * プルダウン選択値に応じて「隣のセル」の背景色が変化する Excel を生成する。
 *
 * <p>仕様:
 * <ul>
 *   <li>偶数列（選択列）でプルダウンから色を選択できる。</li>
 *   <li>選択した色に応じて、右隣の奇数列（メモ列）の背景色が変化する（条件付き書式）。</li>
 *   <li>その色付きセルに文字列を入力すると、背景色は無し（白）に戻る。</li>
 *   <li>列は 100 列（選択列+メモ列のペア × 50）。行は 10,000 行。</li>
 * </ul>
 *
 * <p>大量セル（最大 100 × 10,000 = 100万セル）になるため SXSSF でストリーミング生成する。
 * 条件付き書式とデータ入力規則はシート単位の範囲指定で 1 度だけ定義するため、行数に
 * 比例した肥大化は起きない。
 */
@RestController
@RequestMapping("/api/excel")
public class ExcelConditionalController {

    static final int TOTAL_COLUMNS = 100;    // 選択列(50) + メモ列(50)
    static final int PAIRS         = TOTAL_COLUMNS / 2;
    static final int DATA_ROWS     = 10_000;  // データ行数（ヘッダー除く）

    /**
     * プルダウンの選択肢と、それに対応するメモ列の背景色。
     * {表示値, 背景色}
     */
    private static final Object[][] COLOR_OPTIONS = {
        {"赤", IndexedColors.RED},
        {"黄", IndexedColors.YELLOW},
        {"緑", IndexedColors.LIGHT_GREEN},
        {"青", IndexedColors.SKY_BLUE},
    };

    @GetMapping("/download-conditional")
    public ResponseEntity<byte[]> downloadConditional() throws IOException {
        byte[] bytes;
        try (SXSSFWorkbook wb = new SXSSFWorkbook(100);   // メモリ上は直近 100 行のみ保持
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = wb.createSheet("色変化デモ");
            writeSheet(wb, sheet);

            wb.write(out);
            bytes = out.toByteArray();
        }

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.set(HttpHeaders.CONTENT_DISPOSITION,
            "attachment; filename=\"conditional_color.xlsx\"");
        return ResponseEntity.ok().headers(headers).body(bytes);
    }

    // ── シート本体 ───────────────────────────────────────────────────────────────

    private void writeSheet(Workbook wb, Sheet sheet) {
        sheet.setDefaultColumnWidth(14);

        CellStyle headerStyle = buildHeaderStyle(wb);
        CellStyle selectStyle = buildSelectStyle(wb);
        CellStyle memoStyle   = buildMemoStyle(wb);

        // ── ヘッダー行（POI row 0）──
        Row hRow = sheet.createRow(0);
        for (int k = 0; k < PAIRS; k++) {
            int selCol  = 2 * k;       // 選択列
            int memoCol = 2 * k + 1;   // メモ列（色が変化する隣のセル）
            Cell sc = hRow.createCell(selCol);
            sc.setCellValue("選択" + (k + 1));
            sc.setCellStyle(headerStyle);
            Cell mc = hRow.createCell(memoCol);
            mc.setCellValue("メモ" + (k + 1));
            mc.setCellStyle(headerStyle);
        }

        // ── データ行（POI rows 1..DATA_ROWS）──
        for (int i = 0; i < DATA_ROWS; i++) {
            Row r = sheet.createRow(i + 1);
            for (int k = 0; k < PAIRS; k++) {
                int selCol  = 2 * k;
                int memoCol = 2 * k + 1;

                // 選択列: 色をプルダウンで選べる。初期表示用に色を循環セット。
                Cell sc = r.createCell(selCol);
                sc.setCellValue((String) COLOR_OPTIONS[(i + k) % COLOR_OPTIONS.length][0]);
                sc.setCellStyle(selectStyle);

                // メモ列: 5 行に 1 行はサンプル文字列を入れて「入力済みは色なし」を例示。
                // それ以外は空欄のままにし、条件付き書式で背景色が付く。
                if (i % 5 == 4) {
                    Cell mc = r.createCell(memoCol);
                    mc.setCellValue("入力済み");
                    mc.setCellStyle(memoStyle);
                }
            }
        }

        // ── プルダウン（全選択列に一括設定）──
        addDataValidation(sheet);

        // ── 条件付き書式（全メモ列に一括設定）──
        addConditionalFormatting(sheet);
    }

    /**
     * 全選択列（偶数列）に同一のプルダウン（明示リスト）を設定する。
     * 複数の列範囲を 1 つの DataValidation にまとめることで XML の肥大化を防ぐ。
     */
    private void addDataValidation(Sheet sheet) {
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();

        String[] values = new String[COLOR_OPTIONS.length];
        for (int i = 0; i < COLOR_OPTIONS.length; i++) values[i] = (String) COLOR_OPTIONS[i][0];

        DataValidationConstraint constraint = dvHelper.createExplicitListConstraint(values);

        CellRangeAddressList range = new CellRangeAddressList();
        for (int k = 0; k < PAIRS; k++) {
            int selCol = 2 * k;
            range.addCellRangeAddress(1, selCol, DATA_ROWS, selCol);
        }

        DataValidation dv = dvHelper.createValidation(constraint, range);
        dv.setShowErrorBox(true);
        dv.setErrorStyle(DataValidation.ErrorStyle.WARNING);
        dv.createErrorBox("入力エラー", "プルダウンから色を選択してください。");
        dv.setShowPromptBox(true);
        dv.createPromptBox("色選択", "▼ から色を選択すると右隣のセルの背景色が変わります");
        sheet.addValidationData(dv);
    }

    /**
     * 全メモ列（奇数列）に条件付き書式を設定する。
     *
     * <p>数式は列を固定しない相対参照で書く（{@code AND(A2="赤",B2="")}）。これにより同一の
     * 条件付き書式を全メモ列の範囲（B,D,F,…）へ適用したとき、Excel が各範囲の左上セルを基準に
     * 相対的に評価し、「左隣の選択列＝色名」「自セルが空」という関係が各ペアで成立する。
     *
     * <p>各ルールの意味: 左隣の選択列が指定色 かつ 自セルが空 のとき、その背景色を塗る。
     * 自セルに文字列が入力されると {@code B2=""} が偽になり、どのルールにも一致せず背景色なしになる。
     */
    private void addConditionalFormatting(Sheet sheet) {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();

        // 適用範囲: 全メモ列（奇数列）。先頭は B 列（index 1）= 数式の基準セル B2。
        CellRangeAddress[] regions = new CellRangeAddress[PAIRS];
        for (int k = 0; k < PAIRS; k++) {
            int memoCol = 2 * k + 1;
            regions[k] = new CellRangeAddress(1, DATA_ROWS, memoCol, memoCol);
        }

        ConditionalFormattingRule[] rules = new ConditionalFormattingRule[COLOR_OPTIONS.length];
        for (int i = 0; i < COLOR_OPTIONS.length; i++) {
            String value       = (String) COLOR_OPTIONS[i][0];
            IndexedColors color = (IndexedColors) COLOR_OPTIONS[i][1];

            // 基準セル B2 に対する相対数式（左隣 A2 が color かつ B2 が空）
            ConditionalFormattingRule rule =
                scf.createConditionalFormattingRule("AND(A2=\"" + value + "\",B2=\"\")");
            PatternFormatting pf = rule.createPatternFormatting();
            pf.setFillBackgroundColor(color.getIndex());
            pf.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            rules[i] = rule;
        }

        scf.addConditionalFormatting(regions, rules);
    }

    // ── スタイル ─────────────────────────────────────────────────────────────────

    private CellStyle buildHeaderStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true); f.setColor(IndexedColors.WHITE.getIndex()); s.setFont(f);
        s.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        border(s);
        return s;
    }

    /** 選択列（プルダウン）スタイル。薄い黄色で入力可能を示す。 */
    private CellStyle buildSelectStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        border(s);
        return s;
    }

    /** メモ列のうち文字列入力済みセルのスタイル（枠線のみ・背景なし）。 */
    private CellStyle buildMemoStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setAlignment(HorizontalAlignment.LEFT);
        border(s);
        return s;
    }

    private void border(CellStyle s) {
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
    }
}
