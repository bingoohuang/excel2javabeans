package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColAlign;
import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColStyle;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import lombok.var;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Slf4j
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
        List<ExcelBeanField> beanFields = new ArrayList<>(declaredFields.length);

        for (val field : declaredFields) {
            processField(field, beanFields);
        }

        return filterTitledFileds(beanFields);
    }

    private List<ExcelBeanField> filterTitledFileds(List<ExcelBeanField> beanFields) {
        val titledBeanFields = beanFields.stream().filter(x -> x.hasTitle()).collect(Collectors.toList());
        if (titledBeanFields.isEmpty()) return beanFields;

        if (log.isDebugEnabled()) {
            val untitledFields = beanFields.stream().filter(x -> !x.hasTitle()).collect(Collectors.toList());
            untitledFields.forEach(x -> log.debug("ignore field {} without @ExcelColTitle", x.getFieldName()));
        }

        return titledBeanFields;
    }

    private void processField(Field field, List<ExcelBeanField> fields) {
        if (Modifier.isStatic(field.getModifiers())) return;
        // A synthetic field is a compiler-created field that links a local inner class
        // to a block's local variable or reference type parameter.
        // refer: https://javapapers.com/core-java/java-synthetic-class-method-field/
        if (field.isSynthetic()) return;

        if (field.isAnnotationPresent(ExcelColIgnore.class)) return;

        val fieldName = field.getName();
        if (fieldName.startsWith("$")) return; // ignore un-normal fields like $jacocoData

        val beanField = new ExcelBeanField(beanClass, field, fields.size());
        if (beanField.isElementTypeSupported()) {
            setStyle(field, beanField);
            fields.add(beanField);
        } else {
            log.debug("bean field {} was ignored by unsupported type {}", beanField.getFieldName(), beanField.getElementType());
        }
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
        if (colStyle.align() == ExcelColAlign.LEFT) {
            style.setAlignment(HorizontalAlignment.LEFT);
        } else if (colStyle.align() == ExcelColAlign.CENTER) {
            style.setAlignment(HorizontalAlignment.CENTER);
        } else if (colStyle.align() == ExcelColAlign.RIGHT) {
            style.setAlignment(HorizontalAlignment.RIGHT);
        } else {
            style = null;
        }

        return style;
    }

}
