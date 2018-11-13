package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColAlign;
import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColStyle;
import com.github.bingoohuang.utils.reflect.Fields;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Slf4j
public class ExcelBeanFieldParser {
    private final Sheet sheet;
    private final Field[] declaredFields;

    public ExcelBeanFieldParser(Class<?> beanClass, Sheet sheet) {
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
        val titledFields = beanFields.stream().filter(ExcelBeanField::hasTitle)
                .collect(Collectors.toList());
        if (titledFields.isEmpty()) return beanFields;

        if (log.isDebugEnabled()) {
            beanFields.stream().filter(x -> !x.hasTitle())
                    .forEach(x -> log.debug("ignore field {} without @ExcelColTitle", x.getField()));
        }

        return titledFields;
    }

    private void processField(Field field, List<ExcelBeanField> fields, ReflectAsmCache reflectAsmCache) {
        if (Fields.shouldIgnored(field, ExcelColIgnore.class)) return;

        val bf = new ExcelBeanField(field, fields.size(), reflectAsmCache);
        if (bf.isElementTypeSupported()) {
            setStyle(field, bf);
            fields.add(bf);
        } else {
            log.debug("bean field {} was ignored by unsupported type {}",
                    field, bf.getElementType());
        }
    }


    private void setStyle(Field field, ExcelBeanField beanField) {
        val colStyle = field.getAnnotation(ExcelColStyle.class);
        if (colStyle == null) return;

        beanField.setCellStyle(createAlign(colStyle));
    }

    private CellStyle createAlign(ExcelColStyle colStyle) {
        val style = sheet.getWorkbook().createCellStyle();
        style.setAlignment(convertAlign(colStyle.align(), style.getAlignmentEnum()));
        return style;
    }

    private HorizontalAlignment convertAlign(ExcelColAlign align, HorizontalAlignment defaultAlign) {
        switch (align) {
            case LEFT:
                return HorizontalAlignment.LEFT;
            case CENTER:
                return HorizontalAlignment.CENTER;
            case RIGHT:
                return HorizontalAlignment.RIGHT;
        }

        return defaultAlign;
    }

}
