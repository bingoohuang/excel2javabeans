package com.github.bingoohuang.excel2maps;

import lombok.Value;
import org.apache.commons.lang3.StringUtils;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/15.
 */
@Value public class ColumnDef {
    String title;
    String columnName;
    String ignorePattern;

    public ColumnDef(String title, String columnName) {
        this(title, columnName, null);
    }

    public ColumnDef(String title, String columnName, String ignorePattern) {
        this.title = StringUtils.upperCase(title);
        this.columnName = columnName;
        this.ignorePattern = StringUtils.upperCase(ignorePattern);
    }
}
