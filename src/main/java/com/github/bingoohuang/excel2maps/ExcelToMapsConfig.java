package com.github.bingoohuang.excel2maps;

import com.google.common.collect.Lists;
import lombok.Value;

import java.util.List;

@Value
public class ExcelToMapsConfig {
    List<ColumnDef> columnDefs = Lists.newArrayList();

    public void add(ColumnDef columnDef) {
        columnDefs.add(columnDef);
    }
}
