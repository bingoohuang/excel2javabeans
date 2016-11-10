package com.github.bingoohuang.excel2javabeans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import com.github.bingoohuang.excel2javabeans.annotations.ExcelColumnIgnore;
import com.github.bingoohuang.excel2javabeans.annotations.ExcelColumnTitle;
import com.github.bingoohuang.excel2javabeans.impl.ExcelBeanField;
import com.google.common.collect.Lists;
import lombok.val;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.objenesis.ObjenesisStd;
import org.objenesis.instantiator.ObjectInstantiator;

import java.lang.reflect.Field;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.capitalize;
import static org.apache.commons.lang3.StringUtils.isEmpty;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
public class ExcelToBeans<T> {
    private final Class<T> beanClass;
    private final FieldAccess fieldAccess;
    private final MethodAccess methodAccess;
    private final ObjectInstantiator<T> instantiator;
    private final ExcelBeanField[] beanFields;
    private final boolean hasTitle;

    public ExcelToBeans(Class<T> beanClass) {
        this.beanClass = beanClass;
        this.fieldAccess = FieldAccess.get(beanClass);
        this.methodAccess = MethodAccess.get(beanClass);
        this.instantiator = new ObjenesisStd().getInstantiatorOf(beanClass);
        this.beanFields = parseBeanFields(beanClass);
        this.hasTitle = hasTitle();
    }

    private boolean hasTitle() {
        for (ExcelBeanField beanField : beanFields) {
            if (beanField.hasTitle()) return true;
        }

        return false;
    }

    public List<T> convert(Workbook workbook) {
        List<T> beans = Lists.newArrayList();

        Sheet sheet = workbook.getSheetAt(0);

        val startRowNum = jumpToStartRow(sheet);

        for (int i = startRowNum, ii = sheet.getLastRowNum(); i <= ii; ++i) {
            T o = instantiator.newInstance();

            Row row = sheet.getRow(i);
            for (int j = 0; j < beanFields.length; ++j) {
                ExcelBeanField beanField = beanFields[j];
                Cell cell = row.getCell(beanField.getColumnIndex());

                String cellValue = cell.getStringCellValue();
                if (isEmpty(cellValue)) continue;

                beanFields[j].setFieldValue(fieldAccess, methodAccess, o, cellValue);
            }

            if (o instanceof ExcelRowIgnore) {
                ExcelRowIgnore ignore = (ExcelRowIgnore) o;
                if (ignore.ignoreRow()) continue;
            }

            if (o instanceof ExcelRowReference) {
                ExcelRowReference ref = (ExcelRowReference) o;
                ref.setRowNum(i);
            }

            beans.add(o);
        }

        return beans;
    }

    private int jumpToStartRow(Sheet sheet) {
        int i = sheet.getFirstRowNum();
        if (!hasTitle) return i;

        for (int ii = sheet.getLastRowNum(); i <= ii; ++i) {
            Row row = sheet.getRow(i);

            boolean containsTitle = false;
            for (int j = 0; j < beanFields.length; ++j) {
                ExcelBeanField beanField = beanFields[j];
                if (!beanField.hasTitle()) {
                    beanField.setColumnIndex(j + row.getFirstCellNum());
                } else {
                    if (findColumn(row, beanField)) containsTitle = true;
                }
            }

            if (containsTitle) return i + 1;
        }

        return i;
    }

    private boolean findColumn(Row row, ExcelBeanField beanField) {
        for (int k = row.getFirstCellNum(); k <= row.getLastCellNum(); ++k) {
            Cell cell = row.getCell(k);
            if (cell == null) continue;

            String cellValue = cell.getStringCellValue();
            if (beanField.containTitle(cellValue)) {
                beanField.setColumnIndex(cell.getColumnIndex());
                return true;
            }
        }
        return false;
    }

    private ExcelBeanField[] parseBeanFields(Class<T> beanClass) {
        Field[] declaredFields = beanClass.getDeclaredFields();
        List<ExcelBeanField> fields = Lists.newArrayList();

        for (Field field : declaredFields) {
            val rowIgnore = field.getAnnotation(ExcelColumnIgnore.class);
            if (rowIgnore != null) continue;

            val beanField = new ExcelBeanField();

            beanField.setColumnIndex(fields.size());
            beanField.setName(field.getName());
            beanField.setSetter("set" + capitalize(field.getName()));

            val colTitle = field.getAnnotation(ExcelColumnTitle.class);
            if (colTitle != null) beanField.setTitle(colTitle.value());

            fields.add(beanField);
        }

        return fields.toArray(new ExcelBeanField[0]);
    }
}
