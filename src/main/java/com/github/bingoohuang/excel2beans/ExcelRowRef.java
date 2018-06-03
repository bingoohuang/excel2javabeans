package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import lombok.Getter;
import lombok.Setter;

public class ExcelRowRef implements ExcelRowReferable {
    @ExcelColIgnore
    @Getter @Setter private int rowNum;
    @ExcelColIgnore
    @Getter @Setter private String error;

    @Override public String error() {
        return error;
    }
}
