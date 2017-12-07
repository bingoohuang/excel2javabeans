package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColAlign;
import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColStyle;
import lombok.experimental.var;
import lombok.val;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.List;

public class ExcelBeanFieldParser {
    private final Class<?> beanClass;
    private final Sheet sheet;
    private final Field[] declaredFields;

    public ExcelBeanFieldParser(Class<?> beanClass, Sheet sheet) {
        this.beanClass = beanClass;
        this.sheet = sheet;
        this.declaredFields = beanClass.getDeclaredFields();
    }

    public List<ExcelBeanField> parseBeanFields() {
        val fields = new ArrayList<ExcelBeanField>(declaredFields.length);

        for (val field : declaredFields) {
            processField(field, fields);
        }

        return fields;
    }

    private void processField(Field field, List<ExcelBeanField> fields) {
        if (Modifier.isStatic(field.getModifiers())) return;
        if (Modifier.isTransient(field.getModifiers())) return;

        val rowIgnore = field.getAnnotation(ExcelColIgnore.class);
        if (rowIgnore != null) return;

        val fieldName = field.getName();
        if (fieldName.startsWith("$")) return; // ignore un-normal fields like $jacocoData

        val beanField = new ExcelBeanField(beanClass, field, fields.size());
        setStyle(field, beanField);
        fields.add(beanField);
    }


    private void setStyle(Field field, ExcelBeanField beanField) {
        val colStyle = field.getAnnotation(ExcelColStyle.class);
        if (colStyle == null) return;

        val style = createAlign(colStyle);
        if (style == null) return;

        beanField.setCellStyle(style);
    }

    private CellStyle createAlign(ExcelColStyle colStyle) {
        var style = sheet.getWorkbook().createCellStyle();
        val align = colStyle.align();
        if (align == ExcelColAlign.LEFT) {
            style.setAlignment(HorizontalAlignment.LEFT);
        } else if (align == ExcelColAlign.CENTER) {
            style.setAlignment(HorizontalAlignment.CENTER);
        } else if (align == ExcelColAlign.RIGHT) {
            style.setAlignment(HorizontalAlignment.RIGHT);
        } else {
            style = null;
        }

        return style;
    }

}
