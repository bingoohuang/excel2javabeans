package com.github.bingoohuang.beans2excel.beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.sql.Timestamp;

/**
 * Created by bingoohuang on 2017/3/20.
 */
@Data @AllArgsConstructor
@ExcelSheet(name = "排期")
public class Schedule {
    @ExcelColTitle("日期")
    private Timestamp time;
    @ExcelColTitle("排期数")
    private int schedules;
    @ExcelColTitle("定课数")
    private int subscribes;
    @ExcelColTitle("其中：小班课")
    private int publics;
    @ExcelColTitle("其中：私教课")
    private int privates;
}
