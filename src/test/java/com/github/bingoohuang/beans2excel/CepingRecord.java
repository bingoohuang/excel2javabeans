package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.annotations.ExcelTemplateSheet;
import lombok.Builder;
import lombok.Data;
import lombok.Singular;

import java.util.Map;

@ExcelTemplateSheet(titleRowRef = 3, templateRowRef = 4, templateSheetName = "批量导出") @Builder @Data
public class CepingRecord {
    @ExcelColTitle("项目名称")
    private String itemName;
    @ExcelColTitle("内外部")
    private String source;
    @ExcelColTitle("姓名")
    private String name;
    @ExcelColTitle @Singular
    private Map<String, String> details;
}
