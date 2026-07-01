package com.example.poi.domain;

import java.time.LocalDate;

/**
 * sales テーブルの1行を表すレコード。
 * MyBatis の map-underscore-to-camel-case 設定により
 * sale_date -> saleDate のようにカラムがマッピングされる。
 */
public record Sale(
        Long id,
        LocalDate saleDate,
        String productName,
        String category,
        int quantity,
        int unitPrice
) {
    /** 数量 × 単価 の合計金額。 */
    public long total() {
        return (long) quantity * unitPrice;
    }
}
