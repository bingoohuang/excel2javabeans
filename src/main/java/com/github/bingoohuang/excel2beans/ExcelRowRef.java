package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import lombok.Getter;
import lombok.Setter;

public class ExcelRowRef implements ExcelRowReferable {
    @ExcelColIgnore
    private int rowNum;
    @ExcelColIgnore
    @Getter @Setter private String error;

    @Override public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    @Override public int getRowNum() {
        return rowNum;
    }

    @Override public String error() {
        return error;
    }
}
