package com.github.bingoohuang.excel2javabeans;

import com.github.bingoohuang.excel2javabeans.annotations.ExcelColumnIgnore;
import lombok.Data;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data
public class ExcelRowReference {
    @ExcelColumnIgnore
    private int rowNum;
}
