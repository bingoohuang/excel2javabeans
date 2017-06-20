package com.github.bingoohuang.excel2beans;

import lombok.Builder;
import lombok.Data;

@Data @Builder
public class CellData {
    private String value; // 单元格取值
    private String comment; // 单元格批注
    private String commentAuthor; // 单元格批注
    private int row; // row index
    private int col; // col index
    private int sheetIndex; // sheet index
}
