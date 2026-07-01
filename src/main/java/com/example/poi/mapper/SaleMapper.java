package com.example.poi.mapper;

import com.example.poi.domain.Sale;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Select;

import java.util.List;

/**
 * sales テーブルへのアクセスを行う MyBatis マッパー。
 */
@Mapper
public interface SaleMapper {

    @Select("SELECT id, sale_date, product_name, category, quantity, unit_price "
            + "FROM sales ORDER BY sale_date")
    List<Sale> findAll();
}
