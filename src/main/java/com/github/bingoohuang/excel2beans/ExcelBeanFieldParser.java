package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColAlign;
import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColStyle;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
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

    public List<ExcelBeanField> parseBeanFields(ReflectAsmCache reflectAsmCache) {
        List<ExcelBeanField> beanFields = new ArrayList<>(declaredFields.length);

        for (val field : declaredFields) {
            processField(field, beanFields, reflectAsmCache);
        }

        return filterTitledFields(beanFields);
    }

    private List<ExcelBeanField> filterTitledFields(List<ExcelBeanField> beanFields) {
        val titledFields = beanFields.stream().filter(x -> x.hasTitle()).collect(Collectors.toList());
        if (titledFields.isEmpty()) return beanFields;

        if (log.isDebugEnabled()) {
            beanFields.stream().filter(x -> !x.hasTitle()).forEach(x -> log.debug("ignore field {} without @ExcelColTitle", x.getFieldName()));
        }

        return titledFields;
    }

    private void processField(Field field, List<ExcelBeanField> fields, ReflectAsmCache reflectAsmCache) {
        if (Modifier.isStatic(field.getModifiers())) return;
        // A synthetic field is a compiler-created field that links a local inner class
        // to a block's local variable or reference type parameter.
        // refer: https://javapapers.com/core-java/java-synthetic-class-method-field/
        if (field.isSynthetic()) return;
        if (field.isAnnotationPresent(ExcelColIgnore.class)) return;
        // ignore un-normal fields like $jacocoData
        if (field.getName().startsWith("$")) return;

        val bf = new ExcelBeanField(beanClass, field, fields.size(), reflectAsmCache);
        if (bf.isElementTypeSupported()) {
            setStyle(field, bf);
            fields.add(bf);
        } else {
            log.debug("bean field {} was ignored by unsupported type {}", bf.getFieldName(), bf.getElementType());
        }
    }


    private void setStyle(Field field, ExcelBeanField beanField) {
        val colStyle = field.getAnnotation(ExcelColStyle.class);
        if (colStyle == null) return;

        beanField.setCellStyle(createAlign(colStyle));
    }

    private CellStyle createAlign(ExcelColStyle colStyle) {
        val style = sheet.getWorkbook().createCellStyle();
        val align = convertAlign(colStyle.align());
        if (align != null) style.setAlignment(align);

        return style;

    }

    private HorizontalAlignment convertAlign(ExcelColAlign align) {
        switch (align) {
            case LEFT:
                return HorizontalAlignment.LEFT;
            case CENTER:
                return HorizontalAlignment.CENTER;
            case RIGHT:
                return HorizontalAlignment.RIGHT;
        }

        return null;
    }

}
