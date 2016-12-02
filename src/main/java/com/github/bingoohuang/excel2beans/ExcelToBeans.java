package com.github.bingoohuang.excel2beans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.impl.ExcelBeanField;
import com.google.common.collect.Lists;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.objenesis.ObjenesisStd;
import org.objenesis.instantiator.ObjectInstantiator;

import java.io.InputStream;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.*;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
public class ExcelToBeans<T> {
    private final FieldAccess fieldAccess;
    private final MethodAccess methodAccess;
    private final ObjectInstantiator<T> instantiator;
    private final ExcelBeanField[] beanFields;
    private final boolean hasTitle;
    final DataFormatter cellFormatter = new DataFormatter();

    public ExcelToBeans(Class<T> beanClass) {
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

    @SneakyThrows public List<T> convert(InputStream excelInputStream) {
        val workbook = WorkbookFactory.create(excelInputStream);
        return convert(workbook);
    }

    public List<T> convert(Workbook workbook) {
        List<T> beans = Lists.newArrayList();

        val sheet = workbook.getSheetAt(0);
        val startRowNum = jumpToStartDataRow(sheet);

        for (int i = startRowNum, ii = sheet.getLastRowNum(); i <= ii; ++i) {
            T o = instantiator.newInstance();

            val row = sheet.getRow(i);
            if (row == null) continue;

            int emptyNum = 0;
            for (int j = 0; j < beanFields.length; ++j) {
                val cell = row.getCell(beanFields[j].getColumnIndex());
                val cellValue = getCellValue(cell);
                if (isEmpty(cellValue)) {
                    ++emptyNum;
                } else {
                    beanFields[j].setFieldValue(fieldAccess, methodAccess, o, cellValue);
                }
            }

            if (emptyNum == beanFields.length) continue;

            if (o instanceof ExcelRowIgnorable) {
                val ignore = (ExcelRowIgnorable) o;
                if (ignore.ignoreRow()) continue;
            }

            if (o instanceof ExcelRowRef) {
                val ref = (ExcelRowRef) o;
                ref.setRowNum(i);
            }

            beans.add(o);
        }

        return beans;
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return null;

        val cellValue = cellFormatter.formatCellValue(cell);
        return trim(cellValue);
    }


    private int jumpToStartDataRow(Sheet sheet) {
        int i = sheet.getFirstRowNum();
        if (!hasTitle) return i;

        // try to find the title row
        for (int ii = sheet.getLastRowNum(); i <= ii; ++i) {
            val row = sheet.getRow(i);

            boolean containsTitle = false;
            for (int j = 0; j < beanFields.length; ++j) {
                val beanField = beanFields[j];
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
        for (int k = row.getFirstCellNum(), kk = row.getLastCellNum(); k <= kk; ++k) {
            Cell cell = row.getCell(k);
            if (cell == null) continue;

            val cellValue = cell.getStringCellValue();
            if (beanField.containTitle(cellValue)) {
                beanField.setColumnIndex(cell.getColumnIndex());
                return true;
            }
        }

        return false;
    }

    private ExcelBeanField[] parseBeanFields(Class<T> beanClass) {
        val declaredFields = beanClass.getDeclaredFields();
        List<ExcelBeanField> fields = Lists.newArrayList();

        for (val field : declaredFields) {
            val rowIgnore = field.getAnnotation(ExcelColIgnore.class);
            if (rowIgnore != null) continue;

            val beanField = new ExcelBeanField();

            beanField.setColumnIndex(fields.size());
            beanField.setName(field.getName());
            beanField.setSetter("set" + capitalize(field.getName()));

            val colTitle = field.getAnnotation(ExcelColTitle.class);
            if (colTitle != null) beanField.setTitle(colTitle.value());

            fields.add(beanField);
        }

        return fields.toArray(new ExcelBeanField[0]);
    }
}
