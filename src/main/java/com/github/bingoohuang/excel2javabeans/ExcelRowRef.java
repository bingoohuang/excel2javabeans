package com.github.bingoohuang.excel2javabeans;

import com.github.bingoohuang.excel2javabeans.annotations.ExcelColIgnore;
import lombok.Data;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data
public class ExcelRowRef {
    @ExcelColIgnore private int rowNum;
}
