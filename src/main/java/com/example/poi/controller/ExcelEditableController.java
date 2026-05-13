package com.example.poi.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequestMapping("/api/excel")
public class ExcelEditableController {

    /** 入力行数 */
    private static final int INPUT_ROWS = 30;

    /**
     * 商品マスタ: {商品コード, 商品名, カテゴリ, 標準単価}
     * 売上入力シートのプルダウンと VLOOKUP がこの一覧を参照します。
     */
    private static final Object[][] PRODUCT_MASTER = {
        {"P001", "ノートPC",       "電子機器", 120000},
        {"P002", "デスクトップPC", "電子機器", 200000},
        {"P003", "タブレット",     "電子機器",  80000},
        {"P004", "モニター",       "電子機器",  45000},
        {"P005", "Webカメラ",      "電子機器",  15000},
        {"P006", "スピーカー",     "電子機器",  25000},
        {"P007", "マウス",         "周辺機器",   3500},
        {"P008", "キーボード",     "周辺機器",   8000},
        {"P009", "ヘッドセット",   "周辺機器",  12000},
        {"P010", "USBハブ",        "周辺機器",   2000},
        {"P011", "USBケーブル",    "周辺機器",    800},
        {"P012", "マウスパッド",   "周辺機器",   1500},
        {"P013", "プリンター",     "周辺機器",  35000},
        {"P014", "SSDドライブ",    "記憶媒体",   8500},
        {"P015", "外付けHDD",      "記憶媒体",  12000},
        {"P016", "SDカード",       "記憶媒体",   2000},
    };

    @GetMapping("/download-editable")
    public ResponseEntity<byte[]> downloadEditable() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            // 入力シートを先に作成（開いたとき最初に表示される）
            Sheet inputSheet  = wb.createSheet("売上入力");
            Sheet masterSheet = wb.createSheet("商品マスタ");

            writeMasterSheet(wb, masterSheet);
            writeInputSheet(wb, inputSheet);

            wb.setActiveSheet(0);
            wb.write(out);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.set(HttpHeaders.CONTENT_DISPOSITION,
                "attachment; filename=\"sales_editable.xlsx\"");
            return ResponseEntity.ok().headers(headers).body(out.toByteArray());
        }
    }

    // ── Sheet 2: 商品マスタ ───────────────────────────────────────────────────

    private void writeMasterSheet(Workbook wb, Sheet sheet) {
        sheet.setColumnWidth(0, 3500);  // 商品コード
        sheet.setColumnWidth(1, 6000);  // 商品名
        sheet.setColumnWidth(2, 4000);  // カテゴリ
        sheet.setColumnWidth(3, 5000);  // 標準単価

        // タイトル (POI row 0 = Excel row 1)
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(28);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("商品マスタ（プルダウン参照元）");
        titleCell.setCellStyle(buildTitleStyle(wb, IndexedColors.DARK_TEAL));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

        // ヘッダー (POI row 1 = Excel row 2)
        CellStyle headerStyle = buildHeaderStyle(wb);
        Row hRow = sheet.createRow(1);
        hRow.setHeightInPoints(20);
        String[] hdrs = {"商品コード", "商品名", "カテゴリ", "標準単価（円）"};
        for (int i = 0; i < hdrs.length; i++) strCell(hRow, i, hdrs[i], headerStyle);

        // データ (POI rows 2〜 = Excel rows 3〜)
        CellStyle dataStyle = buildDataStyle(wb);
        CellStyle numStyle  = buildNumStyle(wb);
        for (int i = 0; i < PRODUCT_MASTER.length; i++) {
            Object[] p = PRODUCT_MASTER[i];
            Row r = sheet.createRow(i + 2);
            r.setHeightInPoints(18);
            strCell(r, 0, (String) p[0], dataStyle);
            strCell(r, 1, (String) p[1], dataStyle);
            strCell(r, 2, (String) p[2], dataStyle);
            numCell(r, 3, (int) p[3], numStyle);
        }
    }

    // ── Sheet 1: 売上入力 ─────────────────────────────────────────────────────

    private void writeInputSheet(Workbook wb, Sheet sheet) {
        sheet.setColumnWidth(0, 4500);  // 日付
        sheet.setColumnWidth(1, 6000);  // 商品名
        sheet.setColumnWidth(2, 4000);  // カテゴリ
        sheet.setColumnWidth(3, 3000);  // 数量
        sheet.setColumnWidth(4, 4500);  // 単価
        sheet.setColumnWidth(5, 5000);  // 合計

        CellStyle titleStyle      = buildTitleStyle(wb, IndexedColors.ROYAL_BLUE);
        CellStyle noteStyle       = buildNoteStyle(wb);
        CellStyle headerStyle     = buildHeaderStyle(wb);
        CellStyle inputStyle      = buildInputStyle(wb);
        CellStyle inputDateStyle  = buildInputDateStyle(wb);
        CellStyle formulaStyle    = buildFormulaStyle(wb);
        CellStyle formulaNumStyle = buildFormulaNumStyle(wb);
        CellStyle totalLabel      = buildTotalLabelStyle(wb);
        CellStyle totalNum        = buildTotalNumStyle(wb);

        // ── Row 0: タイトル ──
        Row r0 = sheet.createRow(0);
        r0.setHeightInPoints(32);
        Cell tc = r0.createCell(0);
        tc.setCellValue("売上入力シート");
        tc.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

        // ── Row 1: 入力ガイド ──
        Row r1 = sheet.createRow(1);
        r1.setHeightInPoints(20);
        Cell nc = r1.createCell(0);
        nc.setCellValue(
            "【入力方法】 黄色のセルに入力してください。" +
            "商品名はプルダウンから選択すると、カテゴリ・単価が自動設定されます。合計も自動計算されます。");
        nc.setCellStyle(noteStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 5));

        // ── Row 2: ヘッダー ──
        Row r2 = sheet.createRow(2);
        r2.setHeightInPoints(20);
        String[] hdrs = {"日付", "商品名", "カテゴリ", "数量", "単価（円）", "合計（円）"};
        for (int i = 0; i < hdrs.length; i++) strCell(r2, i, hdrs[i], headerStyle);

        // ── Rows 3〜32 (POI): 入力行 ──
        //   POI row 3  = Excel row 4  (firstDataExcelRow)
        //   POI row 32 = Excel row 33 (lastDataExcelRow)
        final int firstDataPOIRow   = 3;
        final int firstDataExcelRow = firstDataPOIRow + 1;   // 4
        final int lastDataExcelRow  = firstDataExcelRow + INPUT_ROWS - 1; // 33

        for (int i = 0; i < INPUT_ROWS; i++) {
            int poi  = firstDataPOIRow + i;
            int xRow = firstDataExcelRow + i;  // Excel 行番号（数式で使用）
            Row r = sheet.createRow(poi);
            r.setHeightInPoints(18);

            // A: 日付（入力・黄色・日付書式）
            r.createCell(0).setCellStyle(inputDateStyle);

            // B: 商品名（入力・黄色・プルダウン）
            r.createCell(1).setCellStyle(inputStyle);

            // C: カテゴリ（VLOOKUP 自動・灰色）
            Cell catCell = r.createCell(2);
            catCell.setCellFormula(
                "IF(B" + xRow + "=\"\",\"\"," +
                "VLOOKUP(B" + xRow + ",'商品マスタ'!$B:$D,2,FALSE))");
            catCell.setCellStyle(formulaStyle);

            // D: 数量（入力・黄色）
            r.createCell(3).setCellStyle(inputStyle);

            // E: 単価（VLOOKUP 自動・灰色）
            Cell priceCell = r.createCell(4);
            priceCell.setCellFormula(
                "IF(B" + xRow + "=\"\",\"\"," +
                "VLOOKUP(B" + xRow + ",'商品マスタ'!$B:$D,3,FALSE))");
            priceCell.setCellStyle(formulaNumStyle);

            // F: 合計（D×E 自動・灰色）
            Cell totalCell = r.createCell(5);
            totalCell.setCellFormula(
                "IF(OR(D" + xRow + "=\"\",E" + xRow + "=\"\"),\"\"," +
                "D" + xRow + "*E" + xRow + ")");
            totalCell.setCellStyle(formulaNumStyle);
        }

        // ── 合計行 ──
        int totalPOI = firstDataPOIRow + INPUT_ROWS;  // 33
        Row tr = sheet.createRow(totalPOI);
        tr.setHeightInPoints(22);
        sheet.addMergedRegion(new CellRangeAddress(totalPOI, totalPOI, 0, 2));
        for (int i = 0; i <= 2; i++) {
            Cell c = tr.createCell(i);
            if (i == 0) c.setCellValue("合計");
            c.setCellStyle(totalLabel);
        }
        Cell sumQty = tr.createCell(3);
        sumQty.setCellFormula(
            "SUM(D" + firstDataExcelRow + ":D" + lastDataExcelRow + ")");
        sumQty.setCellStyle(totalNum);
        tr.createCell(4).setCellStyle(totalLabel);  // 単価列は空
        Cell sumTotal = tr.createCell(5);
        sumTotal.setCellFormula(
            "SUM(F" + firstDataExcelRow + ":F" + lastDataExcelRow + ")");
        sumTotal.setCellStyle(totalNum);

        // ── プルダウン（データ入力規則）──
        //   商品マスタの商品名列: Excel rows 3〜(PRODUCT_MASTER.length+2)
        String listFormula = "'商品マスタ'!$B$3:$B$" + (PRODUCT_MASTER.length + 2);
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint =
            dvHelper.createFormulaListConstraint(listFormula);
        CellRangeAddressList range =
            new CellRangeAddressList(firstDataPOIRow, firstDataPOIRow + INPUT_ROWS - 1, 1, 1);
        DataValidation dv = dvHelper.createValidation(constraint, range);
        dv.setShowErrorBox(true);
        dv.setErrorStyle(DataValidation.ErrorStyle.WARNING);
        dv.createErrorBox("入力エラー", "商品マスタにない商品名です。プルダウンから選択してください。");
        dv.setShowPromptBox(true);
        dv.createPromptBox("商品選択", "▼ からリストを選択してください");
        sheet.addValidationData(dv);
    }

    // ── スタイル生成 ──────────────────────────────────────────────────────────

    private CellStyle buildTitleStyle(Workbook wb, IndexedColors bg) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true); f.setFontHeightInPoints((short) 16);
        f.setColor(IndexedColors.WHITE.getIndex()); s.setFont(f);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setFillForegroundColor(bg.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return s;
    }

    private CellStyle buildNoteStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setColor(IndexedColors.DARK_BLUE.getIndex()); s.setFont(f);
        s.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.LEFT);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        return s;
    }

    private CellStyle buildHeaderStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true); f.setColor(IndexedColors.WHITE.getIndex()); s.setFont(f);
        s.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        border(s, BorderStyle.THIN); return s;
    }

    private CellStyle buildDataStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setAlignment(HorizontalAlignment.CENTER);
        border(s, BorderStyle.THIN); return s;
    }

    private CellStyle buildNumStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0"));
        s.setAlignment(HorizontalAlignment.RIGHT);
        border(s, BorderStyle.THIN); return s;
    }

    /** ユーザー入力セル（黄色） */
    private CellStyle buildInputStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        border(s, BorderStyle.THIN); return s;
    }

    /** 日付入力セル（黄色・日付書式） */
    private CellStyle buildInputDateStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setDataFormat(wb.createDataFormat().getFormat("yyyy/mm/dd"));
        s.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        border(s, BorderStyle.THIN); return s;
    }

    /** 自動計算セル・文字列（灰色） */
    private CellStyle buildFormulaStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        border(s, BorderStyle.THIN); return s;
    }

    /** 自動計算セル・数値（灰色） */
    private CellStyle buildFormulaNumStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0"));
        s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.RIGHT);
        border(s, BorderStyle.THIN); return s;
    }

    private CellStyle buildTotalLabelStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont(); f.setBold(true); s.setFont(f);
        s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        border(s, BorderStyle.MEDIUM); return s;
    }

    private CellStyle buildTotalNumStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont(); f.setBold(true); s.setFont(f);
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0"));
        s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.RIGHT);
        border(s, BorderStyle.MEDIUM); return s;
    }

    private void border(CellStyle s, BorderStyle bs) {
        s.setBorderTop(bs); s.setBorderBottom(bs);
        s.setBorderLeft(bs); s.setBorderRight(bs);
    }

    // ── セル書き込みショートハンド ────────────────────────────────────────────

    private void strCell(Row r, int col, String v, CellStyle s) {
        Cell c = r.createCell(col); c.setCellValue(v); c.setCellStyle(s);
    }

    private void numCell(Row r, int col, long v, CellStyle s) {
        Cell c = r.createCell(col); c.setCellValue(v); c.setCellStyle(s);
    }
}
