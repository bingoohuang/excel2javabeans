package com.github.bingoohuang.beans2excel.beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColStyle;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import lombok.AllArgsConstructor;
import lombok.Data;

import static com.github.bingoohuang.excel2beans.annotations.ExcelColAlign.*;

/**
 * Created by bingoohuang on 2017/3/20.
 */
@Data
@AllArgsConstructor
@ExcelSheet(name = "会员", headKey = "memberHead")
public class Member {
    @ExcelColTitle("会员总数")
    @ExcelColStyle(align = LEFT)
    private int total;
    @ExcelColTitle("其中：新增")
    @ExcelColStyle(align = CENTER)
    private int fresh;
    @ExcelColTitle("其中：有效")
    @ExcelColStyle(align = RIGHT)
    private int effective;
}
