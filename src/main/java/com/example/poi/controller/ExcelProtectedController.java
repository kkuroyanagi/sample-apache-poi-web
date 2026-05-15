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
public class ExcelProtectedController {

    /** 入力行数 */
    private static final int INPUT_ROWS = 20;

    /**
     * カテゴリマスタ: {カテゴリコード, カテゴリ名, 説明}
     * このデータは保護されたシートで参照専用となります。
     */
    private static final Object[][] CATEGORY_MASTER = {
        {"C001", "電子機器", "PC、タブレット、モニターなどの電子機器"},
        {"C002", "周辺機器", "マウス、キーボード、プリンターなどの周辺機器"},
        {"C003", "記憶媒体", "SSD、HDD、SDカードなどの記憶媒体"},
        {"C004", "ソフトウェア", "アプリケーション、OS、開発ツールなど"},
        {"C005", "文具用品", "ノート、ペン、ファイルなどの事務用品"},
    };

    /**
     * 商品マスタ: {商品コード, 商品名, カテゴリコード, 標準単価}
     * カテゴリから選択された商品の詳細データです。
     */
    private static final Object[][] PRODUCT_MASTER = {
        {"P001", "ノートPC",           "C001", 120000},
        {"P002", "デスクトップPC",     "C001", 200000},
        {"P003", "タブレット",         "C001",  80000},
        {"P004", "モニター",           "C001",  45000},
        {"P005", "Webカメラ",          "C001",  15000},
        {"P006", "スピーカー",         "C001",  25000},
        {"P007", "マウス",             "C002",   3500},
        {"P008", "キーボード",         "C002",   8000},
        {"P009", "ヘッドセット",       "C002",  12000},
        {"P010", "USBハブ",            "C002",   2000},
        {"P011", "USBケーブル",        "C002",    800},
        {"P012", "マウスパッド",       "C002",   1500},
        {"P013", "プリンター",         "C002",  35000},
        {"P014", "SSDドライブ",        "C003",   8500},
        {"P015", "外付けHDD",          "C003",  12000},
        {"P016", "SDカード",           "C003",   2000},
        {"P017", "オフィススイート",   "C004",  25000},
        {"P018", "開発環境",           "C004",  50000},
        {"P019", "ノート",             "C005",    200},
        {"P020", "ボールペン",         "C005",    150},
    };

    @GetMapping("/download-protected")
    public ResponseEntity<byte[]> downloadProtected() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            // 編集用シートを先に作成（開いたとき最初に表示される）
            Sheet editSheet = wb.createSheet("カテゴリ別売上入力");
            Sheet categoryMasterSheet = wb.createSheet("カテゴリマスタ");
            Sheet productMasterSheet = wb.createSheet("商品マスタ");

            // マスタシートを作成（保護付き）
            writeCategoryMasterSheet(wb, categoryMasterSheet);
            writeProductMasterSheet(wb, productMasterSheet);
            
            // 編集シートを作成
            writeEditSheet(wb, editSheet);

            // マスタシートを保護（パスワードなしで構造保護）
            protectMasterSheets(categoryMasterSheet, productMasterSheet);

            wb.setActiveSheet(0);
            wb.write(out);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.set(HttpHeaders.CONTENT_DISPOSITION,
                "attachment; filename=\"protected_category_sales.xlsx\"");
            return ResponseEntity.ok().headers(headers).body(out.toByteArray());
        }
    }

    // ── Sheet 1: カテゴリ別売上入力 ───────────────────────────────────────────────────

    private void writeEditSheet(Workbook wb, Sheet sheet) {
        sheet.setColumnWidth(0, 4500);  // 日付
        sheet.setColumnWidth(1, 6000);  // カテゴリ
        sheet.setColumnWidth(2, 6000);  // 商品名
        sheet.setColumnWidth(3, 3000);  // 数量
        sheet.setColumnWidth(4, 4500);  // 単価
        sheet.setColumnWidth(5, 5000);  // 合計
        sheet.setColumnWidth(6, 6000);  // 備考

        CellStyle titleStyle      = buildTitleStyle(wb, IndexedColors.DARK_GREEN);
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
        tc.setCellValue("カテゴリ別売上入力シート（マスタ保護）");
        tc.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));

        // ── Row 1: 入力ガイド ──
        Row r1 = sheet.createRow(1);
        r1.setHeightInPoints(20);
        Cell nc = r1.createCell(0);
        nc.setCellValue(
            "【操作方法】 黄色のセルに入力してください。" +
            "カテゴリを選択すると、対応する商品のプルダウンが利用可能になります。" +
            "商品を選択すると単価が自動設定されます。");
        nc.setCellStyle(noteStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));

        // ── Row 2: ヘッダー ──
        Row r2 = sheet.createRow(2);
        r2.setHeightInPoints(20);
        String[] hdrs = {"日付", "カテゴリ", "商品名", "数量", "単価（円）", "合計（円）", "備考"};
        for (int i = 0; i < hdrs.length; i++) strCell(r2, i, hdrs[i], headerStyle);

        // ── Rows 3〜22 (POI): 入力行 ──
        final int firstDataPOIRow   = 3;
        final int firstDataExcelRow = firstDataPOIRow + 1;   // 4
        final int lastDataExcelRow  = firstDataExcelRow + INPUT_ROWS - 1; // 23

        for (int i = 0; i < INPUT_ROWS; i++) {
            int poi  = firstDataPOIRow + i;
            int xRow = firstDataExcelRow + i;  // Excel 行番号（数式で使用）
            Row r = sheet.createRow(poi);
            r.setHeightInPoints(18);

            // A: 日付（入力・黄色・日付書式）
            r.createCell(0).setCellStyle(inputDateStyle);

            // B: カテゴリ（入力・黄色・プルダウン）
            r.createCell(1).setCellStyle(inputStyle);

            // C: 商品名（入力・黄色・プルダウン）
            r.createCell(2).setCellStyle(inputStyle);

            // D: 数量（入力・黄色）
            r.createCell(3).setCellStyle(inputStyle);

            // E: 単価（VLOOKUP 自動・灰色）
            Cell priceCell = r.createCell(4);
            priceCell.setCellFormula(
                "IF(C" + xRow + "=\"\",\"\"," +
                "VLOOKUP(C" + xRow + ",'商品マスタ'!$B:$D,3,FALSE))");
            priceCell.setCellStyle(formulaNumStyle);

            // F: 合計（D×E 自動・灰色）
            Cell totalCell = r.createCell(5);
            totalCell.setCellFormula(
                "IF(OR(D" + xRow + "=\"\",E" + xRow + "=\"\"),\"\"," +
                "D" + xRow + "*E" + xRow + ")");
            totalCell.setCellStyle(formulaNumStyle);

            // G: 備考（入力・黄色）
            r.createCell(6).setCellStyle(inputStyle);
        }

        // ── 合計行 ──
        int totalPOI = firstDataPOIRow + INPUT_ROWS;  // 23
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
        tr.createCell(6).setCellStyle(totalLabel);  // 備考列は空

        // ── プルダウン（データ入力規則）──
        setupDataValidation(sheet, firstDataPOIRow, INPUT_ROWS);
    }

    // ── Sheet 2: カテゴリマスタ（保護） ─────────────────────────────────────────────

    private void writeCategoryMasterSheet(Workbook wb, Sheet sheet) {
        sheet.setColumnWidth(0, 4000);  // カテゴリコード
        sheet.setColumnWidth(1, 5000);  // カテゴリ名
        sheet.setColumnWidth(2, 12000); // 説明

        // タイトル (POI row 0 = Excel row 1)
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(28);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("カテゴリマスタ（保護済み・参照専用）");
        titleCell.setCellStyle(buildTitleStyle(wb, IndexedColors.DARK_RED));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));

        // ヘッダー (POI row 1 = Excel row 2)
        CellStyle headerStyle = buildHeaderStyle(wb);
        Row hRow = sheet.createRow(1);
        hRow.setHeightInPoints(20);
        String[] hdrs = {"カテゴリコード", "カテゴリ名", "説明"};
        for (int i = 0; i < hdrs.length; i++) strCell(hRow, i, hdrs[i], headerStyle);

        // データ (POI rows 2〜 = Excel rows 3〜)
        CellStyle dataStyle = buildDataStyle(wb);
        for (int i = 0; i < CATEGORY_MASTER.length; i++) {
            Object[] c = CATEGORY_MASTER[i];
            Row r = sheet.createRow(i + 2);
            r.setHeightInPoints(18);
            strCell(r, 0, (String) c[0], dataStyle);
            strCell(r, 1, (String) c[1], dataStyle);
            strCell(r, 2, (String) c[2], dataStyle);
        }
    }

    // ── Sheet 3: 商品マスタ（保護） ─────────────────────────────────────────────────

    private void writeProductMasterSheet(Workbook wb, Sheet sheet) {
        sheet.setColumnWidth(0, 3500);  // 商品コード
        sheet.setColumnWidth(1, 6000);  // 商品名
        sheet.setColumnWidth(2, 4000);  // カテゴリコード
        sheet.setColumnWidth(3, 5000);  // 標準単価

        // タイトル (POI row 0 = Excel row 1)
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(28);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("商品マスタ（保護済み・参照専用）");
        titleCell.setCellStyle(buildTitleStyle(wb, IndexedColors.DARK_RED));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

        // ヘッダー (POI row 1 = Excel row 2)
        CellStyle headerStyle = buildHeaderStyle(wb);
        Row hRow = sheet.createRow(1);
        hRow.setHeightInPoints(20);
        String[] hdrs = {"商品コード", "商品名", "カテゴリコード", "標準単価（円）"};
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

    // ── データ入力規則の設定 ────────────────────────────────────────────────────────

    private void setupDataValidation(Sheet sheet, int firstDataPOIRow, int inputRows) {
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();

        // カテゴリ選択のプルダウン
        String categoryListFormula = "'カテゴリマスタ'!$B$3:$B$" + (CATEGORY_MASTER.length + 2);
        DataValidationConstraint categoryConstraint =
            dvHelper.createFormulaListConstraint(categoryListFormula);
        CellRangeAddressList categoryRange =
            new CellRangeAddressList(firstDataPOIRow, firstDataPOIRow + inputRows - 1, 1, 1);
        DataValidation categoryDV = dvHelper.createValidation(categoryConstraint, categoryRange);
        categoryDV.setShowErrorBox(true);
        categoryDV.setErrorStyle(DataValidation.ErrorStyle.WARNING);
        categoryDV.createErrorBox("入力エラー", "カテゴリマスタにないカテゴリです。プルダウンから選択してください。");
        categoryDV.setShowPromptBox(true);
        categoryDV.createPromptBox("カテゴリ選択", "▼ からカテゴリを選択してください");
        sheet.addValidationData(categoryDV);

        // 商品選択のプルダウン（商品マスタの商品名列全体）
        String productListFormula = "'商品マスタ'!$B$3:$B$" + (PRODUCT_MASTER.length + 2);
        DataValidationConstraint productConstraint =
            dvHelper.createFormulaListConstraint(productListFormula);
        CellRangeAddressList productRange =
            new CellRangeAddressList(firstDataPOIRow, firstDataPOIRow + inputRows - 1, 2, 2);
        DataValidation productDV = dvHelper.createValidation(productConstraint, productRange);
        productDV.setShowErrorBox(true);
        productDV.setErrorStyle(DataValidation.ErrorStyle.WARNING);
        productDV.createErrorBox("入力エラー", "商品マスタにない商品名です。プルダウンから選択してください。");
        productDV.setShowPromptBox(true);
        productDV.createPromptBox("商品選択", "▼ から商品を選択してください");
        sheet.addValidationData(productDV);
    }

    // ── シート保護の設定 ────────────────────────────────────────────────────────────

    private void protectMasterSheets(Sheet categorySheet, Sheet productSheet) {
        // カテゴリマスタシートの保護（構造とセルの編集を禁止）
        categorySheet.protectSheet("");  // パスワードなしで保護
        categorySheet.getCTWorksheet().getSheetProtection().setSelectLockedCells(true);
        categorySheet.getCTWorksheet().getSheetProtection().setSelectUnlockedCells(false);
        categorySheet.getCTWorksheet().getSheetProtection().setFormatCells(false);
        categorySheet.getCTWorksheet().getSheetProtection().setFormatColumns(false);
        categorySheet.getCTWorksheet().getSheetProtection().setFormatRows(false);
        categorySheet.getCTWorksheet().getSheetProtection().setInsertColumns(false);
        categorySheet.getCTWorksheet().getSheetProtection().setInsertRows(false);
        categorySheet.getCTWorksheet().getSheetProtection().setInsertHyperlinks(false);
        categorySheet.getCTWorksheet().getSheetProtection().setDeleteColumns(false);
        categorySheet.getCTWorksheet().getSheetProtection().setDeleteRows(false);

        // 商品マスタシートの保護（構造とセルの編集を禁止）
        productSheet.protectSheet("");  // パスワードなしで保護
        productSheet.getCTWorksheet().getSheetProtection().setSelectLockedCells(true);
        productSheet.getCTWorksheet().getSheetProtection().setSelectUnlockedCells(false);
        productSheet.getCTWorksheet().getSheetProtection().setFormatCells(false);
        productSheet.getCTWorksheet().getSheetProtection().setFormatColumns(false);
        productSheet.getCTWorksheet().getSheetProtection().setFormatRows(false);
        productSheet.getCTWorksheet().getSheetProtection().setInsertColumns(false);
        productSheet.getCTWorksheet().getSheetProtection().setInsertRows(false);
        productSheet.getCTWorksheet().getSheetProtection().setInsertHyperlinks(false);
        productSheet.getCTWorksheet().getSheetProtection().setDeleteColumns(false);
        productSheet.getCTWorksheet().getSheetProtection().setDeleteRows(false);
    }

    // ── スタイル生成 ──────────────────────────────────────────────────────────────

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

    // ── セル書き込みショートハンド ────────────────────────────────────────────────

    private void strCell(Row r, int col, String v, CellStyle s) {
        Cell c = r.createCell(col); c.setCellValue(v); c.setCellStyle(s);
    }

    private void numCell(Row r, int col, long v, CellStyle s) {
        Cell c = r.createCell(col); c.setCellValue(v); c.setCellStyle(s);
    }
}