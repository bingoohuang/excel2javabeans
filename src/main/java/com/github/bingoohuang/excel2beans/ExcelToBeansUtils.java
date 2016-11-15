package com.github.bingoohuang.excel2beans;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.InputStream;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@UtilityClass public class ExcelToBeansUtils {
    @SneakyThrows
    public Workbook getClassPathWorkbook(String classPathExcelName) {
        val classLoader = ExcelToBeansUtils.class.getClassLoader();
        @Cleanup val is = classLoader.getResourceAsStream(classPathExcelName);
        return WorkbookFactory.create(is);
    }

    @SneakyThrows
    public InputStream getClassPathInputStream(String classPathExcelName) {
        val classLoader = ExcelToBeansUtils.class.getClassLoader();
        return classLoader.getResourceAsStream(classPathExcelName);
    }
}
