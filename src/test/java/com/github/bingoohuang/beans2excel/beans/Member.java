package com.github.bingoohuang.beans2excel.beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * Created by bingoohuang on 2017/3/20.
 */
@Data @AllArgsConstructor
@ExcelSheet(name = "会员")
public class Member {
    @ExcelColTitle("会员总数")
    private int total;
    @ExcelColTitle("其中：新增")
    private int fresh;
    @ExcelColTitle("其中：有效")
    private int effective;
}
