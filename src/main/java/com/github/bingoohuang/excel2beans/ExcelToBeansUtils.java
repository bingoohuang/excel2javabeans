package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.Lists;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.InputStream;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.capitalize;

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

    public static ExcelBeanField[] parseBeanFields(Class<?> beanClass) {
        val declaredFields = beanClass.getDeclaredFields();
        List<ExcelBeanField> fields = Lists.newArrayList();

        for (val field : declaredFields) {
            val rowIgnore = field.getAnnotation(ExcelColIgnore.class);
            if (rowIgnore != null) continue;

            val beanField = new ExcelBeanField();

            beanField.setColumnIndex(fields.size());
            beanField.setName(field.getName());
            beanField.setSetter("set" + capitalize(field.getName()));
            beanField.setGetter("get" + capitalize(field.getName()));

            val colTitle = field.getAnnotation(ExcelColTitle.class);
            if (colTitle != null) beanField.setTitle(colTitle.value());

            fields.add(beanField);
        }

        return fields.toArray(new ExcelBeanField[0]);
    }
}
