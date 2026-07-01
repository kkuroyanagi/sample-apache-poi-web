-- 売上テーブル（ExcelDbController の Excel 出力データソース）
CREATE TABLE IF NOT EXISTS sales (
    id           SERIAL PRIMARY KEY,
    sale_date    DATE NOT NULL,
    product_name VARCHAR(100) NOT NULL,
    category     VARCHAR(50)  NOT NULL,
    quantity     INT          NOT NULL,
    unit_price   INT          NOT NULL
);
