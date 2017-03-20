package com.github.bingoohuang.beans2excel.beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.sql.Timestamp;

/**
 * Created by bingoohuang on 2017/3/20.
 */
@Data
@AllArgsConstructor
@ExcelSheet(name = "定课")
public class Subscribe {
    @ExcelColTitle("订单日期")
    private Timestamp day;
    @ExcelColTitle("人次")
    private int times;
    @ExcelColTitle("人数")
    private int heads;
}
