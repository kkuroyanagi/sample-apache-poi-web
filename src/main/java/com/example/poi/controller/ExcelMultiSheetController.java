package com.example.poi.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

@RestController
@RequestMapping("/api/excel")
public class ExcelMultiSheetController {

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    private static final Object[][] APRIL_DATA = {
        {LocalDate.of(2025, 4, 1),  "ノートPC",       "電子機器", 3,  120000},
        {LocalDate.of(2025, 4, 3),  "マウス",         "周辺機器", 10, 3500},
        {LocalDate.of(2025, 4, 5),  "キーボード",     "周辺機器", 5,  8000},
        {LocalDate.of(2025, 4, 10), "モニター",       "電子機器", 2,  45000},
        {LocalDate.of(2025, 4, 15), "USBハブ",        "周辺機器", 8,  2000},
        {LocalDate.of(2025, 4, 18), "Webカメラ",      "電子機器", 4,  15000},
        {LocalDate.of(2025, 4, 22), "ヘッドセット",   "周辺機器", 6,  12000},
        {LocalDate.of(2025, 4, 25), "デスクトップPC", "電子機器", 1,  200000},
        {LocalDate.of(2025, 4, 28), "プリンター",     "周辺機器", 2,  35000},
        {LocalDate.of(2025, 4, 30), "SSDドライブ",    "記憶媒体", 15, 8500},
    };

    private static final Object[][] MAY_DATA = {
        {LocalDate.of(2025, 5, 2),  "タブレット",   "電子機器", 3,  80000},
        {LocalDate.of(2025, 5, 5),  "マウス",       "周辺機器", 15, 3500},
        {LocalDate.of(2025, 5, 8),  "キーボード",   "周辺機器", 8,  8000},
        {LocalDate.of(2025, 5, 12), "モニター",     "電子機器", 3,  45000},
        {LocalDate.of(2025, 5, 14), "USBケーブル",  "周辺機器", 20, 800},
        {LocalDate.of(2025, 5, 19), "スピーカー",   "電子機器", 5,  25000},
        {LocalDate.of(2025, 5, 21), "マウスパッド", "周辺機器", 12, 1500},
        {LocalDate.of(2025, 5, 26), "ノートPC",     "電子機器", 2,  120000},
        {LocalDate.of(2025, 5, 28), "外付けHDD",    "記憶媒体", 7,  12000},
        {LocalDate.of(2025, 5, 30), "SDカード",     "記憶媒体", 25, 2000},
    };

    @GetMapping("/download-multi")
    public ResponseEntity<byte[]> downloadMultiSheet() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Styles s = new Styles(wb);

            // 5 シートを順番に作成
            writeSalesSheet(wb.createSheet("4月 売上データ"), APRIL_DATA, "売上レポート（2025年4月）", s);
            writeSalesSheet(wb.createSheet("5月 売上データ"), MAY_DATA,   "売上レポート（2025年5月）", s);
            writeCategorySheet(wb.createSheet("カテゴリ別集計"), s);
            writeComparisonSheet(wb.createSheet("月次比較"), s);
            writeSummarySheet(wb.createSheet("サマリー"), s);

            wb.write(out);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.set(HttpHeaders.CONTENT_DISPOSITION,
                "attachment; filename=\"sales_5sheets.xlsx\"");
            return ResponseEntity.ok().headers(headers).body(out.toByteArray());
        }
    }

    // ── Sheet 1 & 2: 月別売上データ ──────────────────────────────────────────

    private static final String[] SALES_HEADERS =
        {"日付", "商品名", "カテゴリ", "数量", "単価（円）", "合計（円）"};

    private void writeSalesSheet(Sheet sheet, Object[][] data, String title, Styles s) {
        int[] widths = {4000, 6000, 4000, 3000, 4500, 4500};
        for (int i = 0; i < widths.length; i++) sheet.setColumnWidth(i, widths[i]);

        mergedTitle(sheet, title, s.title, 5);

        Row hRow = sheet.createRow(1);
        hRow.setHeightInPoints(20);
        for (int i = 0; i < SALES_HEADERS.length; i++) sCell(hRow, i, SALES_HEADERS[i], s.header);

        long grand = 0;
        for (int i = 0; i < data.length; i++) {
            Object[] d = data[i];
            Row r = sheet.createRow(i + 2);
            r.setHeightInPoints(18);
            sCell(r, 0, ((LocalDate) d[0]).format(DATE_FMT), s.data);
            sCell(r, 1, (String) d[1], s.data);
            sCell(r, 2, (String) d[2], s.data);
            int qty = (int) d[3], price = (int) d[4];
            long total = (long) qty * price;
            grand += total;
            nCell(r, 3, qty, s.num);
            nCell(r, 4, price, s.num);
            nCell(r, 5, total, s.num);
        }

        int ti = data.length + 2;
        Row tr = sheet.createRow(ti);
        tr.setHeightInPoints(20);
        sheet.addMergedRegion(new CellRangeAddress(ti, ti, 0, 4));
        for (int i = 0; i <= 4; i++) {
            Cell c = tr.createCell(i);
            if (i == 0) c.setCellValue("合計");
            c.setCellStyle(s.totalLabel);
        }
        nCell(tr, 5, grand, s.totalNum);
    }

    // ── Sheet 3: カテゴリ別集計 ───────────────────────────────────────────────

    private void writeCategorySheet(Sheet sheet, Styles s) {
        int[] widths = {5000, 3000, 5500, 5500, 5500};
        for (int i = 0; i < widths.length; i++) sheet.setColumnWidth(i, widths[i]);

        mergedTitle(sheet, "カテゴリ別集計（2025年4〜5月）", s.title, 4);

        Row hRow = sheet.createRow(1);
        hRow.setHeightInPoints(20);
        String[] hdrs = {"カテゴリ", "件数", "4月合計（円）", "5月合計（円）", "総合計（円）"};
        for (int i = 0; i < hdrs.length; i++) sCell(hRow, i, hdrs[i], s.header);

        Map<String, long[]> aprAgg = aggregate(APRIL_DATA);
        Map<String, long[]> mayAgg = aggregate(MAY_DATA);
        Set<String> cats = mergedKeys(aprAgg, mayAgg);

        int ri = 2;
        long gCount = 0, gApr = 0, gMay = 0;
        for (String cat : cats) {
            long[] a = aprAgg.getOrDefault(cat, new long[2]);
            long[] m = mayAgg.getOrDefault(cat, new long[2]);
            Row r = sheet.createRow(ri++);
            r.setHeightInPoints(18);
            sCell(r, 0, cat, s.data);
            nCell(r, 1, a[0] + m[0], s.num);
            nCell(r, 2, a[1], s.num);
            nCell(r, 3, m[1], s.num);
            nCell(r, 4, a[1] + m[1], s.num);
            gCount += a[0] + m[0]; gApr += a[1]; gMay += m[1];
        }
        Row tr = sheet.createRow(ri);
        tr.setHeightInPoints(20);
        sCell(tr, 0, "合計", s.totalLabel);
        nCell(tr, 1, gCount,        s.totalNum);
        nCell(tr, 2, gApr,          s.totalNum);
        nCell(tr, 3, gMay,          s.totalNum);
        nCell(tr, 4, gApr + gMay,   s.totalNum);
    }

    // ── Sheet 4: 月次比較 ─────────────────────────────────────────────────────

    private void writeComparisonSheet(Sheet sheet, Styles s) {
        int[] widths = {5000, 5500, 5500, 5500, 4000};
        for (int i = 0; i < widths.length; i++) sheet.setColumnWidth(i, widths[i]);

        mergedTitle(sheet, "月次比較（カテゴリ別）", s.title, 4);

        Row hRow = sheet.createRow(1);
        hRow.setHeightInPoints(20);
        String[] hdrs = {"カテゴリ", "4月合計（円）", "5月合計（円）", "増減（円）", "増減率"};
        for (int i = 0; i < hdrs.length; i++) sCell(hRow, i, hdrs[i], s.header);

        Map<String, long[]> aprAgg = aggregate(APRIL_DATA);
        Map<String, long[]> mayAgg = aggregate(MAY_DATA);

        int ri = 2;
        for (String cat : mergedKeys(aprAgg, mayAgg)) {
            long apr  = aprAgg.getOrDefault(cat, new long[2])[1];
            long may  = mayAgg.getOrDefault(cat, new long[2])[1];
            long diff = may - apr;
            double rate = apr == 0 ? 0.0 : (double) diff / apr;

            Row r = sheet.createRow(ri++);
            r.setHeightInPoints(18);
            sCell(r, 0, cat, s.data);
            nCell(r, 1, apr, s.num);
            nCell(r, 2, may, s.num);
            nCell(r, 3, diff, diff > 0 ? s.numPos : diff < 0 ? s.numNeg : s.num);
            dCell(r, 4, rate, diff > 0 ? s.pctPos : diff < 0 ? s.pctNeg : s.pct);
        }
    }

    // ── Sheet 5: サマリー ─────────────────────────────────────────────────────

    private void writeSummarySheet(Sheet sheet, Styles s) {
        sheet.setColumnWidth(0, 7000);
        sheet.setColumnWidth(1, 8000);

        mergedTitle(sheet, "売上サマリー（2025年4〜5月）", s.title, 1);

        long aprTotal = totalAmount(APRIL_DATA);
        long mayTotal = totalAmount(MAY_DATA);
        long grand    = aprTotal + mayTotal;
        long diff     = mayTotal - aprTotal;
        double rate   = (double) diff / aprTotal;

        int ri = 2;
        kvRow(sheet, ri++, "集計期間",       "2025/04/01 〜 2025/05/31", s);
        kvRow(sheet, ri++, "総取引件数",     (APRIL_DATA.length + MAY_DATA.length) + " 件", s);
        sheet.createRow(ri++).setHeightInPoints(10); // 空白行
        kvNumRow(sheet, ri++, "4月 売上合計",   aprTotal, s);
        kvNumRow(sheet, ri++, "5月 売上合計",   mayTotal, s);
        kvNumRow(sheet, ri++, "2ヶ月 売上合計", grand, s);
        kvRow(sheet, ri++, "前月比",
            String.format("%+,.0f 円（%+.1f%%）", (double) diff, rate * 100), s);
        sheet.createRow(ri++).setHeightInPoints(10); // 空白行
        kvRow(sheet, ri++, "4月 最多売上商品", topProduct(APRIL_DATA), s);
        kvRow(sheet, ri++, "5月 最多売上商品", topProduct(MAY_DATA), s);
        kvRow(sheet, ri++, "最多売上カテゴリ", topCategory(APRIL_DATA, MAY_DATA), s);
    }

    private void kvRow(Sheet sheet, int ri, String label, String value, Styles s) {
        Row r = sheet.createRow(ri);
        r.setHeightInPoints(22);
        sCell(r, 0, label, s.summaryLabel);
        sCell(r, 1, value, s.summaryValue);
    }

    private void kvNumRow(Sheet sheet, int ri, String label, long value, Styles s) {
        Row r = sheet.createRow(ri);
        r.setHeightInPoints(22);
        sCell(r, 0, label, s.summaryLabel);
        nCell(r, 1, value, s.summaryCurrency);
    }

    // ── 集計ヘルパー ──────────────────────────────────────────────────────────

    /** カテゴリ → [件数, 合計金額] */
    private Map<String, long[]> aggregate(Object[][] data) {
        Map<String, long[]> map = new LinkedHashMap<>();
        for (Object[] row : data) {
            long[] agg = map.computeIfAbsent((String) row[2], k -> new long[2]);
            agg[0]++;
            agg[1] += (long)(int) row[3] * (int) row[4];
        }
        return map;
    }

    private Set<String> mergedKeys(Map<String, long[]> a, Map<String, long[]> b) {
        Set<String> s = new LinkedHashSet<>(a.keySet());
        s.addAll(b.keySet());
        return s;
    }

    private long totalAmount(Object[][] data) {
        long t = 0;
        for (Object[] row : data) t += (long)(int) row[3] * (int) row[4];
        return t;
    }

    private String topProduct(Object[][] data) {
        String top = ""; long max = 0;
        for (Object[] row : data) {
            long amt = (long)(int) row[3] * (int) row[4];
            if (amt > max) { max = amt; top = (String) row[1]; }
        }
        return top;
    }

    @SafeVarargs
    private String topCategory(Object[][]... datasets) {
        Map<String, Long> totals = new LinkedHashMap<>();
        for (Object[][] data : datasets)
            for (Object[] row : data)
                totals.merge((String) row[2], (long)(int) row[3] * (int) row[4], Long::sum);
        return totals.entrySet().stream()
            .max(Map.Entry.comparingByValue())
            .map(Map.Entry::getKey).orElse("");
    }

    // ── セル書き込みショートハンド ────────────────────────────────────────────

    private void mergedTitle(Sheet sheet, String title, CellStyle style, int lastCol) {
        Row r = sheet.createRow(0);
        r.setHeightInPoints(32);
        Cell c = r.createCell(0);
        c.setCellValue(title);
        c.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastCol));
    }

    private void sCell(Row r, int col, String v, CellStyle s) {
        Cell c = r.createCell(col); c.setCellValue(v); c.setCellStyle(s);
    }

    private void nCell(Row r, int col, long v, CellStyle s) {
        Cell c = r.createCell(col); c.setCellValue(v); c.setCellStyle(s);
    }

    private void dCell(Row r, int col, double v, CellStyle s) {
        Cell c = r.createCell(col); c.setCellValue(v); c.setCellStyle(s);
    }

    // ── スタイル ──────────────────────────────────────────────────────────────

    private static class Styles {
        // 共通
        final CellStyle title, header, data, num, pct, totalLabel, totalNum;
        // 増減色分け（月次比較シート用）
        final CellStyle numPos, numNeg, pctPos, pctNeg;
        // サマリーシート用
        final CellStyle summaryLabel, summaryValue, summaryCurrency;

        Styles(Workbook wb) {
            DataFormat fmt = wb.createDataFormat();
            title      = buildTitle(wb);
            header     = buildHeader(wb);
            data       = buildData(wb);
            num        = buildNum(wb, fmt);
            pct        = buildPct(wb, fmt, (short) -1);
            totalLabel = buildTotalLabel(wb);
            totalNum   = buildTotalNum(wb, fmt);
            numPos     = buildColorNum(wb, fmt, IndexedColors.DARK_GREEN);
            numNeg     = buildColorNum(wb, fmt, IndexedColors.RED);
            pctPos     = buildPct(wb, fmt, IndexedColors.DARK_GREEN.getIndex());
            pctNeg     = buildPct(wb, fmt, IndexedColors.RED.getIndex());
            summaryLabel    = buildSummaryLabel(wb);
            summaryValue    = buildSummaryValue(wb);
            summaryCurrency = buildSummaryCurrency(wb, fmt);
        }

        private static CellStyle buildTitle(Workbook wb) {
            CellStyle s = wb.createCellStyle();
            Font f = wb.createFont();
            f.setBold(true); f.setFontHeightInPoints((short) 16);
            f.setColor(IndexedColors.WHITE.getIndex()); s.setFont(f);
            s.setAlignment(HorizontalAlignment.CENTER);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            s.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return s;
        }

        private static CellStyle buildHeader(Workbook wb) {
            CellStyle s = wb.createCellStyle();
            Font f = wb.createFont();
            f.setBold(true); f.setColor(IndexedColors.WHITE.getIndex()); s.setFont(f);
            s.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            s.setAlignment(HorizontalAlignment.CENTER);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            border(s, BorderStyle.THIN); return s;
        }

        private static CellStyle buildData(Workbook wb) {
            CellStyle s = wb.createCellStyle();
            s.setAlignment(HorizontalAlignment.CENTER);
            border(s, BorderStyle.THIN); return s;
        }

        private static CellStyle buildNum(Workbook wb, DataFormat fmt) {
            CellStyle s = wb.createCellStyle();
            s.setDataFormat(fmt.getFormat("#,##0"));
            s.setAlignment(HorizontalAlignment.RIGHT);
            border(s, BorderStyle.THIN); return s;
        }

        private static CellStyle buildPct(Workbook wb, DataFormat fmt, short color) {
            CellStyle s = wb.createCellStyle();
            if (color >= 0) {
                Font f = wb.createFont(); f.setBold(true); f.setColor(color); s.setFont(f);
            }
            s.setDataFormat(fmt.getFormat("0.0%"));
            s.setAlignment(HorizontalAlignment.RIGHT);
            border(s, BorderStyle.THIN); return s;
        }

        private static CellStyle buildTotalLabel(Workbook wb) {
            CellStyle s = wb.createCellStyle();
            Font f = wb.createFont(); f.setBold(true); s.setFont(f);
            s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            s.setAlignment(HorizontalAlignment.CENTER);
            border(s, BorderStyle.MEDIUM); return s;
        }

        private static CellStyle buildTotalNum(Workbook wb, DataFormat fmt) {
            CellStyle s = wb.createCellStyle();
            Font f = wb.createFont(); f.setBold(true); s.setFont(f);
            s.setDataFormat(fmt.getFormat("#,##0"));
            s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            s.setAlignment(HorizontalAlignment.RIGHT);
            border(s, BorderStyle.MEDIUM); return s;
        }

        private static CellStyle buildColorNum(Workbook wb, DataFormat fmt, IndexedColors color) {
            CellStyle s = wb.createCellStyle();
            Font f = wb.createFont(); f.setBold(true); f.setColor(color.getIndex()); s.setFont(f);
            s.setDataFormat(fmt.getFormat("#,##0"));
            s.setAlignment(HorizontalAlignment.RIGHT);
            border(s, BorderStyle.THIN); return s;
        }

        private static CellStyle buildSummaryLabel(Workbook wb) {
            CellStyle s = wb.createCellStyle();
            Font f = wb.createFont(); f.setBold(true); s.setFont(f);
            s.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            s.setAlignment(HorizontalAlignment.LEFT);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            border(s, BorderStyle.THIN); return s;
        }

        private static CellStyle buildSummaryValue(Workbook wb) {
            CellStyle s = wb.createCellStyle();
            s.setAlignment(HorizontalAlignment.LEFT);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            border(s, BorderStyle.THIN); return s;
        }

        private static CellStyle buildSummaryCurrency(Workbook wb, DataFormat fmt) {
            CellStyle s = wb.createCellStyle();
            Font f = wb.createFont(); f.setBold(true); f.setFontHeightInPoints((short) 12); s.setFont(f);
            s.setDataFormat(fmt.getFormat("¥#,##0"));
            s.setAlignment(HorizontalAlignment.RIGHT);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            border(s, BorderStyle.THIN); return s;
        }

        private static void border(CellStyle s, BorderStyle bs) {
            s.setBorderTop(bs); s.setBorderBottom(bs);
            s.setBorderLeft(bs); s.setBorderRight(bs);
        }
    }
}
