package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColStyle;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.Lists;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import lombok.val;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.List;

import static com.github.bingoohuang.excel2beans.annotations.ExcelColAlign.CENTER;
import static com.github.bingoohuang.excel2beans.annotations.ExcelColAlign.LEFT;
import static com.github.bingoohuang.excel2beans.annotations.ExcelColAlign.RIGHT;
import static org.apache.commons.lang3.StringUtils.capitalize;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@UtilityClass
public class ExcelToBeansUtils {
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

    public static ExcelBeanField[] parseBeanFields(Class<?> beanClass, Sheet sheet) {
        val declaredFields = beanClass.getDeclaredFields();
        List<ExcelBeanField> fields = Lists.newArrayList();

        for (val field : declaredFields) {
            val rowIgnore = field.getAnnotation(ExcelColIgnore.class);
            if (rowIgnore != null) {
                continue;
            }

            String fieldName = field.getName();
            if (fieldName.startsWith("$")) { // ignore un-normal fields like $jacocoData
                continue;
            }

            val beanField = new ExcelBeanField();

            beanField.setColumnIndex(fields.size());
            beanField.setName(fieldName);
            beanField.setSetter("set" + capitalize(fieldName));
            beanField.setGetter("get" + capitalize(fieldName));

            val colTitle = field.getAnnotation(ExcelColTitle.class);
            if (colTitle != null) {
                beanField.setTitle(colTitle.value());
            }

            val colStyle = field.getAnnotation(ExcelColStyle.class);
            if (colStyle != null) {
                CellStyle style = setAlign(sheet, colStyle);
                if (style != null) {
                    beanField.setCellStyle(style);
                }
            }

            fields.add(beanField);
        }

        return fields.toArray(new ExcelBeanField[0]);
    }

    private static CellStyle setAlign(Sheet sheet, ExcelColStyle colStyle) {
        val style = sheet.getWorkbook().createCellStyle();
        if (colStyle.align() == LEFT) {
            style.setAlignment(HorizontalAlignment.LEFT);
            return style;
        } else if (colStyle.align() == CENTER) {
            style.setAlignment(HorizontalAlignment.CENTER);
            return style;
        } else if (colStyle.align() == RIGHT) {
            style.setAlignment(HorizontalAlignment.RIGHT);
            return style;
        } else {
            return null;
        }
    }

    @SneakyThrows
    public static void download(HttpServletResponse response, Workbook workbook, String fileName) {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-disposition", "attachment; filename=" + fileName);
        @Cleanup val out = response.getOutputStream();
        workbook.write(out);
        workbook.close();
    }
}
