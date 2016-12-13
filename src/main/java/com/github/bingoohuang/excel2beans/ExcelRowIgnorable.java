package com.github.bingoohuang.excel2beans;

/**
 * Excel row ignoring interface for the mapping javabean.
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
public interface ExcelRowIgnorable {
    /**
     * to flag out where this row should be ignored.
     * @return ignore current row or not.
     */
    boolean ignoreRow();
}
