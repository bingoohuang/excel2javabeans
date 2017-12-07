package com.github.bingoohuang.excel2beans;

/**
 * bean reference to related excel row.
 */
public interface ExcelRowReferable {
    void setRowNum(int rowNum);

    int getRowNum();

    String error();
}
