package com.github.bingoohuang.excel2javabeans;

import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@UtilityClass
public class ExcelToBeansUtils {
    @SneakyThrows
    public Workbook getClassPathWorkbook(String classPathExcelName) {
        val classLoader = ExcelToBeansUtils.class.getClassLoader();
        val is = classLoader.getResourceAsStream(classPathExcelName);
        return WorkbookFactory.create(is);
    }
}
