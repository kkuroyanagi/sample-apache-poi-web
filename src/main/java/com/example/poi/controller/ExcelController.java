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

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    private static final Object[][] SALES_DATA = {
            {LocalDate.of(2025, 4, 1),  "ノートPC",      "電子機器", 3,  120000},
            {LocalDate.of(2025, 4, 3),  "マウス",        "周辺機器", 10, 3500},
            {LocalDate.of(2025, 4, 5),  "キーボード",    "周辺機器", 5,  8000},
            {LocalDate.of(2025, 4, 10), "モニター",      "電子機器", 2,  45000},
            {LocalDate.of(2025, 4, 15), "USBハブ",       "周辺機器", 8,  2000},
            {LocalDate.of(2025, 4, 18), "Webカメラ",     "電子機器", 4,  15000},
            {LocalDate.of(2025, 4, 22), "ヘッドセット",  "周辺機器", 6,  12000},
            {LocalDate.of(2025, 4, 25), "デスクトップPC","電子機器", 1,  200000},
            {LocalDate.of(2025, 4, 28), "プリンター",    "周辺機器", 2,  35000},
            {LocalDate.of(2025, 4, 30), "SSDドライブ",   "記憶媒体", 15, 8500},
    };

    @GetMapping("/download")
    public ResponseEntity<byte[]> downloadExcel() throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.createSheet("売上レポート");
            sheet.setColumnWidth(0, 4000);
            sheet.setColumnWidth(1, 6000);
            sheet.setColumnWidth(2, 4000);
            sheet.setColumnWidth(3, 3000);
            sheet.setColumnWidth(4, 4500);
            sheet.setColumnWidth(5, 4500);

            CellStyle titleStyle  = createTitleStyle(workbook);
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle dataStyle   = createDataStyle(workbook);
            CellStyle numStyle    = createNumberStyle(workbook);
            CellStyle totalLabel  = createTotalLabelStyle(workbook);
            CellStyle totalNum    = createTotalNumberStyle(workbook);

            // タイトル行
            Row titleRow = sheet.createRow(0);
            titleRow.setHeightInPoints(32);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue("売上レポート（2025年4月）");
            titleCell.setCellStyle(titleStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

            // ヘッダー行
            Row headerRow = sheet.createRow(1);
            headerRow.setHeightInPoints(20);
            String[] headers = {"日付", "商品名", "カテゴリ", "数量", "単価（円）", "合計（円）"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // データ行
            DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy/MM/dd");
            long grandTotal = 0;
            for (int i = 0; i < SALES_DATA.length; i++) {
                Object[] row = SALES_DATA[i];
                Row dataRow = sheet.createRow(i + 2);
                dataRow.setHeightInPoints(18);

                Cell c0 = dataRow.createCell(0);
                c0.setCellValue(((LocalDate) row[0]).format(fmt));
                c0.setCellStyle(dataStyle);

                Cell c1 = dataRow.createCell(1);
                c1.setCellValue((String) row[1]);
                c1.setCellStyle(dataStyle);

                Cell c2 = dataRow.createCell(2);
                c2.setCellValue((String) row[2]);
                c2.setCellStyle(dataStyle);

                int qty = (int) row[3];
                int price = (int) row[4];
                long total = (long) qty * price;
                grandTotal += total;

                Cell c3 = dataRow.createCell(3);
                c3.setCellValue(qty);
                c3.setCellStyle(numStyle);

                Cell c4 = dataRow.createCell(4);
                c4.setCellValue(price);
                c4.setCellStyle(numStyle);

                Cell c5 = dataRow.createCell(5);
                c5.setCellValue(total);
                c5.setCellStyle(numStyle);
            }

            // 合計行
            int totalRowIdx = SALES_DATA.length + 2;
            Row totalRow = sheet.createRow(totalRowIdx);
            totalRow.setHeightInPoints(20);
            sheet.addMergedRegion(new CellRangeAddress(totalRowIdx, totalRowIdx, 0, 4));
            for (int i = 0; i <= 4; i++) {
                Cell c = totalRow.createCell(i);
                if (i == 0) c.setCellValue("合計");
                c.setCellStyle(totalLabel);
            }
            Cell grandTotalCell = totalRow.createCell(5);
            grandTotalCell.setCellValue(grandTotal);
            grandTotalCell.setCellStyle(totalNum);

            workbook.write(out);

            HttpHeaders responseHeaders = new HttpHeaders();
            responseHeaders.setContentType(MediaType.parseMediaType(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            responseHeaders.set(HttpHeaders.CONTENT_DISPOSITION,
                    "attachment; filename=\"sales_report.xlsx\"");

            return ResponseEntity.ok().headers(responseHeaders).body(out.toByteArray());
        }
    }

    private CellStyle createTitleStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short) 16);
        f.setColor(IndexedColors.WHITE.getIndex());
        s.setFont(f);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return s;
    }

    private CellStyle createHeaderStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setColor(IndexedColors.WHITE.getIndex());
        s.setFont(f);
        s.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        setBorder(s, BorderStyle.THIN);
        return s;
    }

    private CellStyle createDataStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setAlignment(HorizontalAlignment.CENTER);
        setBorder(s, BorderStyle.THIN);
        return s;
    }

    private CellStyle createNumberStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0"));
        s.setAlignment(HorizontalAlignment.RIGHT);
        setBorder(s, BorderStyle.THIN);
        return s;
    }

    private CellStyle createTotalLabelStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        s.setFont(f);
        s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        setBorder(s, BorderStyle.MEDIUM);
        return s;
    }

    private CellStyle createTotalNumberStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        s.setFont(f);
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0"));
        s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.RIGHT);
        setBorder(s, BorderStyle.MEDIUM);
        return s;
    }

    private void setBorder(CellStyle s, BorderStyle bs) {
        s.setBorderTop(bs);
        s.setBorderBottom(bs);
        s.setBorderLeft(bs);
        s.setBorderRight(bs);
    }
}
