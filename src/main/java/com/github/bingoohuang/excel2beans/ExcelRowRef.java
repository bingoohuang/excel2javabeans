package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import lombok.Data;

/**
 * bean reference to related excel row.
 */
@Data
public class ExcelRowRef {
    @ExcelColIgnore private int rowNum;

    public String error() {
        return null;
    }
}
